"""组卷：把老师粘贴的题干 → 一份能直接印的卷子（题号、分值、答题卡、答案卷）。

只做 v0.1 承诺的事：**不接题库、不生成题目**。老师从旧卷、备课资料里粘贴文字，
这里负责"排好、编好号、算好分、配好答题卡"。题目本身的对错、教学法判断不归它管。

解析策略（宁可少猜，也不要把老师刚粘的一段话切碎）：
- 大题：行首 一、/二、… 或 Ⅰ./Ⅱ./ 或 "第一部分"
- 小题：行首 1. / 1、/ 1)，若整段没有阿拉伯编号，则认 （1）（2）
- 选项：行首 A. A、 A） 或 A．，可一行多个
- 答案：【答案】/答案：/参考答案： 后面的内容
- 不属于以上任何一条的行，追加到当前小题的题干里（换行保留）
"""

import copy as copy_module
import os
import re

import docx
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_ROW_HEIGHT_RULE
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Mm, Pt, Twips

from .docx_build import build_docx, column_width_twips, fit_table
from .errors import EngineError, require
from .paper import PAPERS, resolve as resolve_paper
from .paths import default_output
from .preflight import preflight

CN_NUM = "一二三四五六七八九十"
ROMAN = "ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ"
RE_SECTION_PART = re.compile(
    r"^\s*(?P<label>第\s*(?P<num>[" + CN_NUM + r"]{1,3})\s*部分\s*[、.．:：]?)\s*(?P<rest>.*)$")
RE_CN_SECTION = re.compile(
    r"^\s*(?P<label>(?P<num>[" + CN_NUM + r"]{1,3})\s*[、.．:：])\s*(?P<rest>.*)$")
RE_ROMAN_SECTION = re.compile(
    r"^\s*(?P<label>(?P<num>[" + ROMAN + r"]{1,3})\s*[、.．:：])\s*(?P<rest>.*)$")
RE_BARE_SECTION = re.compile(
    r"^\s*(?P<label>(?P<num>[" + CN_NUM + r"]{1,3})[\s\u3000])\s*(?P<rest>.*)$")
RE_EMBEDDED_Q = re.compile(r"^(?P<title>[^\d]{0,24}?)\s*(?P<num>\d{1,3})\s*(?P<sep>[、.．)）])\s*(?P<stem>\S.*)$")
# 行首小数不是题号：「0.1 mol/L 的盐酸」里的 0. 会把题干首字吃掉（实测「1 mol/L…」）
RE_ARABIC_Q = re.compile(r"^\s*(\d{1,3})\s*(?:[、)）]|\.(?!\d)|．(?!\d))\s*(.*)$")
RE_PAREN_Q = re.compile(r"^\s*[（(]\s*(\d{1,3})\s*[)）]\s*(.*)$")
# 粘来的说明行（不是题干也不是选项）：「阅读 B 篇，选出最佳选项。」「根据短文内容判断正误。」
RE_INSTRUCTION_LINE = re.compile(
    r"^(?:请|阅读|读|根据|看|听|观察|完成|选出|选择|判断|回答|填写|write|read|choose)"
    r"[^。]{0,60}(?:篇|短文|对话|材料|图片|问题|选项|正误|填空|句子|句|词|句子)?[。\.!！]?$",
    re.IGNORECASE)
RE_OPTION = re.compile(r"^\s*([A-Ha-h])\s*[、.．)）]\s*(.*)$")
RE_SCORE = re.compile(r"[（(]\s*(\d+(?:\.\d+)?)\s*分\s*[)）]")
RE_INLINE_SCORE = re.compile(r"^(.*?)[（(]\s*(\d+(?:\.\d+)?)\s*分\s*[)）]\s*$")
RE_ANSWER = re.compile(r"^\s*(?:【\s*答案\s*】|【\s*参考答案\s*】|参考答案\s*[:：]?|答案\s*[:：]?)\s*(\S.*)$")
RE_ANSWER_RANGE = re.compile(r"^\s*(\d{1,3})\s*[-—–~～至]\s*(\d{1,3})\s*[:：]?\s*([A-Ha-h][A-Ha-h\s]*)$")
# 卷尾那一整段答案：「【答案】1—5 BCBDA 6—10 ABABB…」「参考答案：46. minutes 47. hospital」
RE_ANSWER_HEAD = re.compile(r"^\s*(?:【\s*(?:参考)?答案\s*】|(?:参考)?答案\s*[:：]?)\s*(.*)$")
RE_ANSWER_RANGE_LOOSE = re.compile(r"(\d{1,3})\s*[-—–~～至]\s*(\d{1,3})\s*[:：]?\s*([A-Ha-h][A-Ha-h\s,，、]*)")
RE_ANSWER_ITEM = re.compile(r"(?:^|[\s；;，,。])(\d{1,3})\s*[.、．)）:：]\s*([^\n]*?)(?=(?:\s+\d{1,3}\s*[.、．)）:：])|[\n]|$)")
RE_DECLARED_TOTAL = re.compile(r"满分\s*(\d+(?:\.\d+)?)\s*分")
# 大题行（一、/第一部分/Ⅰ.）：整卷满分绝不会写在这里
RE_BIG_LINE = re.compile(r"^\s*(?:第\s*[" + CN_NUM + r"]{1,3}\s*部分|[" + CN_NUM + ROMAN + r"]{1,3}\s*[、.．:：])")
RE_MINUTES = re.compile(r"时间\s*(\d{1,3})\s*分钟")
RE_OPTION_SPLIT = re.compile(r"(?<![A-Za-z0-9])([A-Ha-h])\s*[、.．)）]\s*")
# 阅读/完形材料的分篇标签：单独一行写 "B"、"Passage B"、"阅读 B 篇"。
# 老师在材料后面跟一串短文，以前这些短文会被拼进上一道题的题干，印出来学生分不清
# 材料归哪一题（实测第 30 题题干里塞进了整篇 B 篇）。
RE_MATERIAL_LABEL = re.compile(
    r"^\s*(?:(?:Passage|Reading|短文|篇章|材料|阅读)\s*)?([A-H])\s*"
    r"(?:篇|则)?\s*(?:[（(][^）)]{0,20}[)）])?\s*[.．、:：]?\s*$", re.IGNORECASE)
RE_ANALYSIS = re.compile(r"^\s*(?:【\s*解析\s*】|解析\s*[:：]|【\s*分析\s*】)\s*(.*)$")
# 材料段的抬头：「材料：」「【材料】」「短文：」——后面的行属于这篇短文，不算上一题的题干续行。
# （题库组卷没材料就是"裸题"：学生看到 "From the passage…" 却找不到 passage）
# rest 里可能直接跟着短文：「材料：Passage B Tom is a boy…」。早先只认「材料：」独占一行，
# 带正文的写法会掉到"上一题续行"的兜底里，整篇 B 篇短文被塞进第 55 题的题干。
RE_MATERIAL_HEAD = re.compile(
    r"^\s*(?:【\s*(?:材料|短文|语篇|篇章|阅读材料)\s*】|(?:材料|短文|语篇|篇章|阅读材料)\s*[:：])\s*(?P<rest>.*)$")
RE_CJK = re.compile(r"[\u4e00-\u9fff]")
# 大题行里的"分"信息有好几种写法，全都要能读：满分 20 分 / 每小题 1 分 / 共 20 小题 / （20 分）
RE_PAPER_COUNT = re.compile(r"共\s*(\d{1,3})\s*小?题")
RE_EACH_SCORE = re.compile(r"每\s*小?题\s*(\d+(?:\.\d+)?)\s*分")
RE_FULL_SCORE = re.compile(r"满分\s*(\d+(?:\.\d+)?)\s*分")
# 「共 10 分」也常写在大题标题里（和「共 10 小题」不是一个意思，别混）
RE_SECTION_TOTAL = re.compile(r"共\s*(\d+(?:\.\d+)?)\s*分")
# 「（共 20 小题，每小题 1 分，满分 20 分）」这种括号整块算分值说明，不是标题的一部分
SCORE_CLAUSE = re.compile(r"[（(][^（()）]{0,60}?(?:\d+(?:\.\d+)?\s*分|\d+\s*小?题)[^（()）]{0,60}?[)）]")
RE_INLINE_ANALYSIS = re.compile(r"^\s*(?:【\s*解析\s*】|【\s*分析\s*】|解析\s*[:：])?\s*(.*)$")
RE_INLINE_ANALYSIS_SPLIT = re.compile(r"【\s*解析\s*】|【\s*分析\s*】|解析\s*[:：]|分析\s*[:：]")


def _roman_to_int(text):
    table = {ch: i + 1 for i, ch in enumerate(ROMAN)}
    return table.get(text, 0)


def _cn_to_int(text):
    digits = {ch: i + 1 for i, ch in enumerate("一二三四五六七八九")}
    if text == "十":
        return 10
    if text.startswith("十"):
        return 10 + digits.get(text[1], 0)
    if "十" in text:
        head, _, tail = text.partition("十")
        return digits.get(head, 0) * 10 + (digits.get(tail, 0) if tail else 0)
    return digits.get(text, 0)


def _strip_score_clauses(text):
    """去掉标题里的分值说明：「听力（共 20 小题，每小题 1 分，满分 20 分）」→「听力」。"""
    cleaned = SCORE_CLAUSE.sub("", text or "")
    cleaned = RE_SCORE.sub("", cleaned)
    return cleaned.strip(" 　")


def _looks_like_title(rest):
    """大题标题长得像标题吗？教师的正文句子不能被当成大题吞掉。"""
    text = _strip_score_clauses(rest)
    if not text:
        return True
    if re.search(r"[。？！!?；;]", text):
        return False
    return len(text) <= 20


def _split_embedded_question(rest):
    """大题行里夹着第一题：「选择题1. 下列字音…」→ ("选择题", "1. 下列字音…")。

    老师从旧卷里复制时，"一、选择题" 和第一题常常在同一行，不切开第一题就整行蒸发。
    """
    text = (rest or "").strip()
    if not text:
        return "", None
    match = RE_EMBEDDED_Q.match(text)
    if not match:
        return text, None
    title = match.group("title").strip(" 　、.．:：")
    stem = match.group("stem").strip()
    if not stem:
        return text, None
    if title and (len(RE_SCORE.sub("", title).strip()) > 16 or re.search(r"[。？！!?；;]$", title)):
        return text, None
    return title, match.group("num") + match.group("sep") + " " + stem


def _paper_declared_total(text):
    """整卷满分只认卷首那一行。

    大题标题里也常写「（共 5 小题，每小题 2 分，满分 10 分）」——那不是整卷满分；
    早先直接全文搜第一个「满分 X 分」，会把它当整卷满分，于是一份 100 分的卷子
    永远报「卷子上写的满分是 10 分」。
    """
    if not text:
        return None
    for raw in text.splitlines()[:10]:
        line = raw.strip()
        if not line:
            continue
        if RE_PAPER_COUNT.search(line) or RE_BIG_LINE.match(line):
            continue
        match = RE_DECLARED_TOTAL.search(line)
        if match:
            return float(match.group(1))
    return None


def _section_declared(rest):
    """这一大题写在标题里的总分：满分 X 分 > 共 N 小题 × 每小题 X 分 > （X 分）。"""
    text = rest or ""
    full = RE_FULL_SCORE.search(text)
    if full:
        return float(full.group(1))
    # 「（每题 2 分，共 10 分）」：先看那个总数，不能拿「2 分」当整节满分
    total = RE_SECTION_TOTAL.search(text)
    if total:
        return float(total.group(1))
    count = RE_PAPER_COUNT.search(text)
    each = RE_EACH_SCORE.search(text)
    if count and each:
        return round(int(count.group(1)) * float(each.group(1)), 2)
    if each:
        return float(each.group(1))
    score = RE_SCORE.search(text)
    return float(score.group(1)) if score else None


def _section_each_score(rest):
    each = RE_EACH_SCORE.search(rest or "")
    return float(each.group(1)) if each else None


def _section_number(text):
    """返回 (编号, 剩余标题, 原编号写法)；不是大题行返回 (0, None, "")。"""
    for pattern, convert, bare in (
        (RE_SECTION_PART, _cn_to_int, False),
        (RE_CN_SECTION, _cn_to_int, False),
        (RE_ROMAN_SECTION, _roman_to_int, False),
        (RE_BARE_SECTION, _cn_to_int, True),
    ):
        match = pattern.match(text)
        if not match:
            continue
        number = convert(match.group("num"))
        if not number:
            continue
        rest = (match.group("rest") or "").strip()
        if not _looks_like_title(rest):
            return 0, None, ""
        if bare and len(RE_SCORE.sub("", rest).strip()) > 12:
            return 0, None, ""
        return number, rest, match.group("label").strip(" 　")
    return 0, None, ""


RE_SUBHEAD = re.compile(r"^\s*[（(]\s*([一二三四五六七八九十]+)\s*[)）]\s*(.+)$")


def _split_sections(text):
    """粗切大题：只要认出编号就切；没认出来的一律算第一节。"""
    lines = [line.rstrip() for line in (text or "").replace("\r\n", "\n").replace("\r", "\n").split("\n")]
    sections = []
    current = {"title": "试题", "label": "", "declaredScore": None, "lines": []}
    for line in lines:
        subhead = RE_SUBHEAD.match(line)
        if subhead and not RE_ARABIC_Q.match(line) and _looks_like_title(subhead.group(2)):
            # 「（一）文言文阅读（10 分）」是卷内小节：单独成节，题号从 1 重新起
            rest = subhead.group(2)
            if current["lines"] or sections:
                sections.append(current)
            current = {
                "title": _strip_score_clauses(rest) or subhead.group(1),
                "label": "（" + subhead.group(1) + "）",
                "declaredScore": _section_declared(rest),
                "perQuestion": _section_each_score(rest),
                "lines": [],
            }
            continue
        number, rest, label = _section_number(line)
        if number and not RE_OPTION.match(line):
            if current["lines"] or sections:
                sections.append(current)
            title, embedded = _split_embedded_question(rest)
            current = {
                "title": title or ("第" + str(number) + "部分"),
                "label": label,
                "declaredScore": _section_declared(rest),
                "perQuestion": _section_each_score(rest),
                "lines": [embedded] if embedded else [],
            }
            continue
        current["lines"].append(line)
    sections.append(current)
    return [section for section in sections if any(line.strip() for line in section["lines"])]


RE_OPTION_HEAD = re.compile(r"^\s*([A-Ha-h])\s*[、.．)）]\s*")


def _options_only(text):
    """整行就是一组选项（完形填空那种「21. A. small  B. beautiful …」）时返回选项表。

    以前这种行被当成题干整行塞进 stem、options 为空：B 卷没有选项可打乱，
    印出来和 A 卷逐字相同，卷上也看不出哪题该填哪个词。
    """
    if not text or not RE_OPTION_HEAD.match(text):
        return []
    parts = _split_option_line(text)
    return parts if len(parts) >= 2 else []


def _split_option_line(line):
    """一行里塞了多个选项时切开；只认 A–H。

    也要认粘在一起的「A)甲B)乙」（字母前是中文，不能要求空格），
    但字母前是英文/数字时不切，免得把 "A book" 里的 A 当选项。
    """
    parts = []
    matches = list(RE_OPTION_SPLIT.finditer(line))
    for index, match in enumerate(matches):
        start = match.end()
        end = matches[index + 1].start() if index + 1 < len(matches) else len(line)
        text = line[start:end].strip(" 　")
        if text:
            parts.append((match.group(1).upper(), text))
    if parts:
        return parts
    match = RE_OPTION.match(line)
    return [(match.group(1).upper(), match.group(2).strip())] if match else []


def _split_inline_questions(line):
    """一行里并排两道题时切开；不是这种形状就返回空列表。

    判据：第一处题号必须在行首，且后面的题号严格递增——「3.14 是圆周率」这种不会被切开。
    """
    matches = list(re.finditer(r"(?:^|(?<=[\s）)。；;]))(\d{1,3})\s*[、.．)）]\s*", line or ""))
    if len(matches) < 2 or matches[0].start() != 0:
        return []
    numbers = [int(match.group(1)) for match in matches]
    for previous, current_number in zip(numbers, numbers[1:]):
        if current_number != previous + 1:
            return []
    chunks = []
    for index, match in enumerate(matches):
        start = match.end()
        end = matches[index + 1].start() if index + 1 < len(matches) else len(line)
        text = line[start:end].strip(" 　")
        if not text:
            return []
        chunks.append((numbers[index], text))
    return chunks


def _question_style(lines):
    """这一节用哪种小题编号：阿拉伯数字优先，其次（1）。"""
    if any(RE_ARABIC_Q.match(line) for line in lines):
        return "arabic"
    if any(RE_PAREN_Q.match(line) for line in lines):
        return "paren"
    return "none"


def _parse_section(section):
    """返回 (preamble, questions, materials)。

    materials 是阅读/完形那种"材料段"：分篇标签（单独一行写 B / Passage B）后面那几段短文
    单独存起来，按"排在第几道题之前"（after）挂在节上。早先这些行会被拼进上一道题的题干，
    印出来第 30 题的题干里塞着整篇 B 篇短文（老师一眼就看出废了）。
    """
    lines = section["lines"]
    style = _question_style(lines)
    questions = []
    preamble = []
    ranges = []
    current = None
    last_was_option = False
    stray_lines = []
    pending_notes = []
    materials = []
    material = None

    def start_question(number):
        nonlocal current, last_was_option, material
        material = None
        current = {"number": number, "stem": [], "options": [], "answer": None, "analysis": None, "score": None}
        questions.append(current)
        last_was_option = False

    for line in lines:
        if not line.strip():
            # 空行是老师自己的分段：后面的东西不再算"选项续行"
            last_was_option = False
            if material is not None:
                if material["lines"] and material["lines"][-1] != "":
                    material["lines"].append("")
            elif current is not None:
                current["stem"].append("")
            elif preamble and preamble[-1] != "":
                preamble.append("")
            continue

        # 「材料：」抬头：从这里到下一题都是短文
        head = RE_MATERIAL_HEAD.match(line)
        if head:
            rest = (head.group("rest") or "").strip()
            # 「材料：整篇短文」这种写法：抬头后面直接跟着正文时，要判断这行到底是新材料的开头，
            # 还是上一道题的题干续行（化学/物理题里「材料：木炭…」就是题干的一部分）。
            # 判据：没有正在写的题（上一节刚结束）、上一行是选项、或这道题已经收到答案 → 新材料。
            boundary = (current is None or last_was_option
                        or bool(current.get("answer")) or bool(current.get("options")))
            if rest and not boundary:
                current["stem"].append(line)
                last_was_option = False
                continue
            lines_new = [rest] if rest else []
            if pending_notes:
                # 紧挨着材料的那句说明（「阅读 B 篇，选出最佳选项。」）属于这篇短文，放在它前面
                lines_new = list(pending_notes) + lines_new
                pending_notes.clear()
            material = {"after": len(questions), "lines": lines_new}
            materials.append(material)
            current = None
            last_was_option = False
            continue

        # 分篇标签（B / Passage B）：材料另起一段，不再往上一题题干里塞
        if current is not None and RE_MATERIAL_LABEL.match(line):
            lines_new = [line.strip()]
            if pending_notes:
                lines_new = list(pending_notes) + lines_new
                pending_notes.clear()
            material = {"after": len(questions), "lines": lines_new}
            materials.append(material)
            current = None
            last_was_option = False
            continue

        answer_range = RE_ANSWER_RANGE.match(line)
        if answer_range and current is not None:
            ranges.append((int(answer_range.group(1)), int(answer_range.group(2)), answer_range.group(3)))
            last_was_option = False
            continue

        answer = RE_ANSWER.match(line)
        if answer and current is not None:
            # 「【答案】1—5 ABCDA」：这是**一组题**的答案，不是当前这道题的答案。
            # 早先整串塞给当前题，答案页就印成「5. 1—5 ABCDA」，第 1—4 题反而没有答案。
            range_in_answer = RE_ANSWER_RANGE_LOOSE.search((answer.group(1) or "").strip())
            if range_in_answer:
                ranges.append((int(range_in_answer.group(1)), int(range_in_answer.group(2)),
                               range_in_answer.group(3)))
                last_was_option = False
                continue
            # 「答案：A 解析：因为…」在一行里：答案与解析分开存，否则答案页会印两遍解析
            answer_text = (answer.group(1) or "").strip()
            pieces = RE_INLINE_ANALYSIS_SPLIT.split(answer_text, maxsplit=1)
            current["answer"] = pieces[0].strip()
            if len(pieces) > 1 and pieces[1].strip():
                current["analysis"] = pieces[1].strip()
            last_was_option = False
            continue
        analysis = RE_ANALYSIS.match(line)
        if analysis and current is not None:
            current["analysis"] = (analysis.group(1) or "").strip()
            last_was_option = False
            continue

        option = RE_OPTION.match(line)
        if option and current is not None:
            # 老师常把四个选项粘在一行：A. apple  B. banana  C. cat  D. dog
            for label, text in _split_option_line(line):
                current["options"].append({"label": label, "text": text})
            last_was_option = True
            continue

        if style == "arabic":
            match = RE_ARABIC_Q.match(line)
        elif style == "paren":
            match = RE_PAREN_Q.match(line)
        else:
            match = None
        if match:
            # 「3. 甲（2 分）  4. 乙（2 分）」一行两题：切不开就会丢一题
            inline = _split_inline_questions(line)
            if inline:
                for number, stem_text in inline:
                    start_question(number)
                    current["stem"].append(stem_text)
                continue
            start_question(int(match.group(1)))
            stem_text = match.group(2).strip()
            leading_options = _options_only(stem_text)
            if leading_options:
                for label, text in leading_options:
                    current["options"].append({"label": label, "text": text})
                last_was_option = True
                continue
            current["stem"].append(stem_text)
            continue

        if material is not None and current is None:
            material["lines"].append(line.strip())
            last_was_option = False
            continue

        if current is None:
            preamble.append(line)
            last_was_option = False
            continue
        # 续行：紧跟着选项的下一行拼到最后一个选项，其余（含空行之后）都拼题干
        # 但「阅读 B 篇，选出最佳选项。」这类**说明行**不能拼进选项：实测会印成
        # 「D. Chinese books 阅读 B 篇，选出最佳选项。」，学生看到选项里混着指令。
        if last_was_option and current["options"] and RE_INSTRUCTION_LINE.match(line.strip()):
            # 先记着，等下一篇材料的「材料：」出现时放到它前面；真到最后也没有材料，
            # 再退回本节开头并给一条警告。
            pending_notes.append(line.strip())
            last_was_option = False
            continue
        if last_was_option and current["options"]:
            current["options"][-1]["text"] += " " + line.strip()
        else:
            current["stem"].append(line)
            last_was_option = False

    if ranges:
        by_number = {question["number"]: question for question in questions}
        for start, end, letters in ranges:
            compact = re.sub(r"\s+", "", letters).upper()
            for offset, number in enumerate(range(start, end + 1)):
                question = by_number.get(number)
                if question is None or question.get("answer") or offset >= len(compact):
                    continue
                question["answer"] = compact[offset]

    for question in questions:
        stem = "\n".join(question["stem"]).strip()
        match = RE_INLINE_SCORE.search(stem)
        if match and match.group(1).strip():
            stem = match.group(1).strip()
            question["score"] = float(match.group(2))
        question["stem"] = stem
        question.pop("stem_lines", None)
    for item in materials:
        while item["lines"] and not item["lines"][-1].strip():
            item["lines"].pop()
    if pending_notes:
        preamble.extend(pending_notes)
        stray_lines.extend(pending_notes)
    return (preamble, questions,
            [item for item in materials if any(line.strip() for line in item["lines"])],
            stray_lines)


def _answer_block_index(lines):
    """找出卷尾"整段答案"从哪一行开始；找不到返回 None。

    判据（宁可不动，也不能把每道题下面的「【答案】A」误切走）：
      1. 行首是【答案】/【参考答案】/答案：/参考答案：；
      2. 除了标记本身，后面还有像答案的内容（1—5 BCBDA、46. minutes），
         或者标记单独占一行、后面几行能解析出答案。
    """
    for index in range(len(lines) - 1, -1, -1):
        match = RE_ANSWER_HEAD.match(lines[index])
        if not match:
            continue
        rest = (match.group(1) or "").strip()
        # 标记行自己带内容时只看这一行（「【答案】A」是逐题答案，不能把后面的题切走）；
        # 标记单独占一行时，往下看几行找答案。
        candidate = rest if rest else "\n".join(lines[index + 1:]).strip()
        if not candidate:
            continue
        answers = _parse_answer_block(candidate)
        if not answers:
            continue
        # 至少两条（或一条 1—5 区间）才当"整段答案"；单条短答案留给逐题解析
        if len(answers) < 2 and not RE_ANSWER_RANGE_LOOSE.search(candidate):
            continue
        # 这一行后面还有选项行 → 它是「第几题到第几题的答案」的分条答案，不是卷尾答案区。
        # 早先只看这一行像答案就切走后面全部内容：大题后写一句【答案】1—5 ABCDA，
        # 整卷后面的题全被吞掉（实测 11 题 40 分只剩 5 题 10 分，answer-block 还只给一句提醒）。
        tail = lines[index + 1:]
        if any(RE_OPTION.match(line) for line in tail if line.strip()):
            continue
        # 后面还有大题行、且大题下面接着长题干 → 同样不是卷尾答案区
        if any(RE_BIG_LINE.match(line) for line in tail if line.strip()):
            long_stem = False
            for line in tail:
                match_question = RE_ARABIC_Q.match(line)
                if match_question and len(match_question.group(2).strip()) >= 12:
                    long_stem = True
                    break
            if long_stem:
                continue
        return index
    return None


def _parse_answer_block(text):
    """把一段答案文字拆成 {题号: 答案}：1—5 BCBDA、46. minutes 47. hospital 都认。"""
    answers = {}
    body = (text or "").replace("\r", "")
    for match in RE_ANSWER_RANGE_LOOSE.finditer(body):
        start = int(match.group(1))
        end = int(match.group(2))
        letters = re.sub(r"[^A-Ha-h]", "", match.group(3)).upper()
        for offset, number in enumerate(range(start, end + 1)):
            if offset >= len(letters):
                break
            answers.setdefault(number, letters[offset])
    for match in RE_ANSWER_ITEM.finditer(body):
        number = int(match.group(1))
        value = (match.group(2) or "").strip(" 　")
        if value and number not in answers:
            answers[number] = value
    return answers


def _split_answer_block(text):
    """把"题目 + 卷尾答案"切开。返回 (正文, 答案文字或 None)。"""
    lines = [line.rstrip() for line in (text or "").replace("\r\n", "\n").replace("\r", "\n").split("\n")]
    index = _answer_block_index(lines)
    if index is None:
        return text, None
    head = RE_ANSWER_HEAD.match(lines[index])
    rest = (head.group(1) or "").strip() if head else ""
    body_lines = lines[:index]
    answer_lines = ([rest] if rest else []) + lines[index + 1:]
    return "\n".join(body_lines), "\n".join(answer_lines)


def parse_paper(text):
    """把粘贴的文字解析成小节与小题目。返回 {title?, sections, warnings}。"""
    require(isinstance(text, str), "text-not-string", "题目要是一段文字（把题干粘贴进来）", "把题从 Word/题库里复制成纯文字，粘贴到对话框里")
    require(bool(text.strip()), "text-required", "没有收到题目文字（把题干粘贴进来）", "题号写在行首（1. / 2、/（3）），大题写「一、选择题（每题 2 分）」")
    # 老师习惯把答案整段贴在卷尾（【答案】1—5 BCBDA…）：先把它切走，否则会被当成新题印进学生卷
    body_text, answer_text = _split_answer_block(text)
    global_answers = _parse_answer_block(answer_text) if answer_text else {}
    raw_sections = _split_sections(body_text)
    warnings = []
    if answer_text is not None:
        if global_answers:
            warnings.append({
                "code": "answer-block",
                "message": "卷尾那段【答案】我当成答案区了（认出 " + str(len(global_answers))
                           + " 条），不会印进学生卷" + ("" if global_answers else "——一条都没读懂，请检查写法"),
            })
        else:
            warnings.append({
                "code": "answer-block-unparsed",
                "message": "卷尾那段【答案】我没读懂，先没排进去；答案请写成 1—5 BCBDA 或 46. minutes 这样",
            })
    sections = []
    stray_lines = []
    for raw in raw_sections:
        preamble, questions, materials, strays = _parse_section(raw)
        stray_lines.extend(strays)
        title = raw["title"]
        label = raw.get("label") or ""
        declared = raw["declaredScore"]
        if not questions:
            heading_text = (label or "") + (title or "")
            notes = [line for line in preamble if line.strip()]
            if notes and title and title != "试题" and RE_ESSAY.search(heading_text):
                # 作文/书面表达常常只写一段要求、没有题号。早先这种小节被并进上一节，
                # 印出来「二、书面表达」整个标题消失，作文要求跑到上一节第 1 题前面
                # （老师一眼就看出卷子错了）。这里当成「只有要求、没有小题」的写作节留着。
                sections.append({
                    "title": _strip_score_clauses(title) or None,
                    "label": label,
                    "declaredScore": declared,
                    "perQuestion": raw.get("perQuestion"),
                    "notes": notes,
                    "questions": [],
                    "materials": materials,
                    "writing": True,
                })
                continue
            # 整节只有文字（说明、听力指令），并到上一节的尾部说明里
            if sections:
                sections[-1]["notes"].extend([line for line in preamble if line.strip()])
                continue
            title = title if title != "试题" else None
            sections.append({"title": _strip_score_clauses(title) or None, "label": label, "declaredScore": declared,
                             "perQuestion": raw.get("perQuestion"),
                             "notes": [line for line in preamble if line.strip()], "questions": [],
                             "materials": materials})
            continue
        sections.append({
            "title": _strip_score_clauses(title) or None,
            "label": label,
            "declaredScore": declared,
            "perQuestion": raw.get("perQuestion"),
            "notes": [line for line in preamble if line.strip()],
            "questions": questions,
            "materials": materials,
        })

    # 老师粘进来的第一行常常就是卷名（"七年级英语期中测试"），第二行是"满分 100 分，时间 90 分钟"。
    # 卷名拿去当标题，满分/时间行不重复印——出卷时会按实际分值重写一行，两行都在就成了两个分数。
    if sections:
        head = sections[0]
        if head.get("title") in (None, "试题") and head["notes"]:
            first = head["notes"][0].strip()
            if first and len(first) <= 30 and not re.search(r"[。？！!?；;]", first):
                head["title"] = first
                head["notes"] = head["notes"][1:]
        head["notes"] = [line for line in head["notes"]
                         if line.strip() and not RE_DECLARED_TOTAL.search(line)]
    # 卷尾答案按题号回填到小题上（题号下面已有的答案优先）
    if global_answers:
        for section in sections:
            for question in section["questions"]:
                if not question.get("answer") and question["number"] in global_answers:
                    question["answer"] = global_answers[question["number"]]

    minutes = RE_MINUTES.search(body_text)

    total = 0.0
    for section in sections:
        questions = section["questions"]
        if not questions:
            # 写作节没有小题，但老师在标题里写了分（「二、书面表达（10 分）」）：
            # 不算进去的话卷首满分就会比老师说的少十分
            if section.get("writing") and section.get("declaredScore") is not None:
                total += float(section["declaredScore"])
            continue
        declared = section["declaredScore"]
        each = section.get("perQuestion")
        missing = [question for question in questions if question["score"] is None]
        if missing:
            if each is not None:
                # 「每小题 1 分」：逐题分是老师明写的，比均摊准
                for question in missing:
                    question["score"] = each
            elif declared is not None:
                per = round(declared / len(questions), 2)
                for question in missing:
                    question["score"] = per
        actual = round(sum(question["score"] or 0 for question in questions), 2)
        if declared is not None and abs(actual - declared) > 0.01:
            # 老师写的大题分和逐题分对不上：不改他写的数，也不悄悄改总分，直接告诉他
            warnings.append({
                "code": "section-score-mismatch",
                "message": (section["title"] or "某一节") + " 你写的是 " + _fmt_score(declared)
                           + " 分，小题加起来 " + _fmt_score(actual)
                           + " 分（我按小题算；要按大题算就把小题分值改一致）",
                "declared": declared,
                "actual": actual,
            })
        total += actual

    declared_total = _paper_declared_total(body_text)
    if declared_total is not None:
        scored_all = not any(question["score"] is None for section in sections for question in section["questions"])
        if scored_all and total > 0 and abs(total - declared_total) > 0.01:
            warnings.append({
                "code": "total-score-mismatch",
                "level": "warn",
                # 纸面一定按小题算印（否则卷子自己跟自己对不上），所以这句话必须把「印成多少」说清，
                # 并给出两种改法：以前只说「我按小题算」，老师不知道该动哪边。
                "message": "卷子上写的满分是 " + _fmt_score(declared_total) + " 分，小题加起来只有 "
                           + _fmt_score(total) + " 分，卷面就按 " + _fmt_score(total) + " 分印了。"
                           "差的 " + _fmt_score(declared_total - total)
                           + " 分多半是某个大题漏标分值：补上就对了；"
                           "确实按 " + _fmt_score(total) + " 分考，就把卷首的「满分」改成这个数",
                "declared": declared_total,
                "actual": total,
            })

    # 题号：真卷子常常 1–15 跨大题连续；也有的一、二、各自从 1 开始。两种都放过。
    all_questions = [question for section in sections for question in section["questions"]]
    numbers = [question["number"] for question in all_questions]
    global_ok = numbers == list(range(1, len(numbers) + 1))
    per_section_ok = all(
        [question["number"] for question in section["questions"]] == list(range(1, len(section["questions"]) + 1))
        for section in sections if section["questions"]
    )
    if numbers and not global_ok and not per_section_ok:
        warnings.append({
            "code": "renumbered",
            "message": "题号不连续（原本是 " + _number_ranges(numbers) + "），已按 1–" + str(len(numbers)) + " 重排",
            "original": numbers[:30],
        })
        for index, question in enumerate(all_questions):
            question["number"] = index + 1

    if not any(section["questions"] for section in sections):
        warnings.append({"code": "no-questions", "message": "没认出任何小题：题号要写在行首（1. / 1、/ 1)），大题写在行首（一、/ Ⅰ.）"})
    # 完形填空最常见的写法是「26.」后面直接跟选项（空在材料里）：这类题以前题干一直是空的，
    # 印出来只有一个孤零零的「26.」，教师卷里也没有任何讲评线索。补一条填空线，
    # 老师和学生都看得出这题在填第几个空。
    filled_blanks = 0
    for section in sections:
        for question in section["questions"]:
            if not (question.get("stem") or "").strip() and question["options"]:
                question["stem"] = "______"
                filled_blanks += 1
    if filled_blanks:
        warnings.append({
            "code": "blank-stem-filled",
            "message": "有 " + str(filled_blanks) + " 道题只有选项、题干在原句里（完形填空常见），"
                       "已按「题号 ______」排版：老师核对一下空格位置对不对",
        })
    empty_stems = sum(1 for section in sections for question in section["questions"] if not question["stem"])
    if empty_stems:
        warnings.append({"code": "empty-stem", "message": "有 " + str(empty_stems) + " 道题没解析出题干，请检查这几行"})
    if stray_lines:
        # 说明行（「阅读 B 篇，选出最佳选项。」）没法只靠文字判断它属于哪篇材料，
        # 与其粘进上一个选项，不如放到本节开头并明确告诉老师。
        warnings.append({
            "code": "stray-line",
            "message": "这几行看着像说明/指令，没找到该放的位置，已放在所在大题开头：" +
                       "；".join(stray_lines[:3]) + ("…" if len(stray_lines) > 3 else ""),
        })
    return {"sections": sections, "warnings": warnings, "totalScore": round(total, 2),
            "declaredTotalScore": declared_total,
            "minutes": int(minutes.group(1)) if minutes else None}


def _number_ranges(numbers):
    """[1,2,3,7,8] → "1–3、7–8"，给老师看的人话。"""
    if not numbers:
        return ""
    parts = []
    start = previous = numbers[0]
    for number in numbers[1:]:
        if number == previous + 1:
            previous = number
            continue
        parts.append(str(start) if start == previous else str(start) + "–" + str(previous))
        start = previous = number
    parts.append(str(start) if start == previous else str(start) + "–" + str(previous))
    return "、".join(parts)


def _strip_score(text):
    """大题标题里带的（20 分）去掉——出卷时按实际题分重新写一遍，不然会出现两个分。"""
    return RE_SCORE.sub("", text or "").strip(" 　、.．")


def _fmt_score(value):
    if value is None:
        return ""
    if abs(value - round(value)) < 1e-6:
        return str(int(round(value)))
    return ("%.1f" % value).rstrip("0").rstrip(".")


def _option_line(options):
    """选项短就排一行（一行看得快），长就一行一个（不挤）。

    判断要看**整行长度**，不能只看单个选项：四个选项各自都短、加起来超过一栏宽度时，
    排成一行会被 Word 折行，最后一个选项（D. …）就被拆成两行了。
    """
    total = 0
    for option in options:
        text = option["text"]
        if len(text) + 4 > 26 or (RE_CJK.search(text) and len(text) > 18):
            return False
        total += len(text) + 4 + 4  # 选项文字 + 题号前缀 + 选项之间的四个空格
    return total <= 70


def _question_blocks(question, default_score=None):
    score = question["score"] if question["score"] is not None else default_score
    head = str(question["number"]) + ". "
    blocks = [{"type": "paragraph", "runs": [{"text": head + question["stem"]}]}]
    if question["options"]:
        if _option_line(question["options"]):
            text = "    ".join(option["label"] + ". " + option["text"] for option in question["options"])
            blocks.append({"type": "paragraph", "runs": [{"text": text}]})
        else:
            for option in question["options"]:
                blocks.append({"type": "paragraph", "runs": [{"text": "    " + option["label"] + ". " + option["text"]}]})
    # 一道题的题干和选项必须待在一起：分栏排版时选项被推到上一栏顶上，
    # 学生会把选项算到别的题头上（实测第 6 题的 A/B/C/D 跑到右栏第一行）
    for block in blocks[:-1]:
        block["keepWithNext"] = True
    blocks[-1]["keepTogether"] = True
    return blocks


def _card_blocks(questions, per_row=10, start=1):
    """答题卡按 10 格分块，每块一张小表（题号行 + 答案行）。

    早先是把所有块塞进**同一张表**：Word 表格列数必须一致，15 道题会出现
    「10 格 + 5 格」两行，第二行右侧空出 6 个虚线格，印出来像漏题；
    印前检查只读第一行题号，还会误报「答题框编号不是 1 到 36 连续」。
    分块建表后每行宽度都等于本块题数。

    题号一律 1..N 连续 —— 卷面里小题按小节重新从 1 起，答题卡上要连续，
    否则格子上出现两个「3」，学生填错行。
    """
    numbers = list(range(start, start + len(questions)))
    blocks = []
    for index in range(0, len(numbers), per_row):
        chunk = numbers[index:index + per_row]
        blocks.append([["题号"] + [str(number) for number in chunk],
                       ["答案"] + ["" for _ in chunk]])
    return blocks


def _answer_blocks(parsed, detail=False):
    """答案页：一串选择题答案连成段（1–5　ABCDA），其它答案一行一题。

    不印分值 —— 分值只在正文印一次，否则印前检查会把同一道题的分算两遍。
    也不用 "11. " 这种写法，免得被当成第 11 题。
    detail=True 时出"教师卷"口径：每题带题干摘要与解析，讲评时不用两手对着看。
    """
    blocks = []
    for section in parsed["sections"]:
        questions = section["questions"]
        if not any(question.get("answer") for question in questions):
            continue
        blocks.append({"type": "paragraph",
                       "runs": [{"text": (section.get("label") or "") + (section["title"] or "试题"), "bold": True}]})
        pending = None  # (首题号, 末题号, 字母串)

        def flush():
            nonlocal pending
            if pending is None:
                return
            start, last, letters = pending
            span = str(start) if start == last else str(start) + "–" + str(last)
            blocks.append({"type": "paragraph", "runs": [{"text": span + "　" + letters}]})
            pending = None

        for question in questions:
            answer = (question.get("answer") or "").strip()
            if not answer:
                continue
            if detail:
                flush()
                stem = re.sub(r"\s+", " ", question.get("stem") or "").strip()
                if len(stem) > 48:
                    # 按词边界截断：硬切 48 字会切出 「…this afte…」这种半个单词
                    cut = stem[:48]
                    boundary = max(cut.rfind(" "), cut.rfind("　"))
                    if boundary >= 30:
                        cut = cut[:boundary]
                    stem = cut.rstrip(" ，,、") + "…"
                text = str(question["number"]) + ". " + answer
                if stem:
                    text += "　" + stem
                if question.get("analysis"):
                    text += "　解析：" + question["analysis"]
                blocks.append({"type": "paragraph", "runs": [{"text": text}]})
                continue
            if re.fullmatch(r"[A-Ha-h]{1,4}", answer) and not question.get("analysis"):
                if pending is not None and question["number"] == pending[1] + 1:
                    pending = (pending[0], question["number"], pending[2] + answer.upper())
                else:
                    flush()
                    pending = (question["number"], question["number"], answer.upper())
                continue
            flush()
            text = str(question["number"]) + "　" + answer
            if question.get("analysis"):
                text += "　（" + question["analysis"] + "）"
            blocks.append({"type": "paragraph", "runs": [{"text": text}]})
        flush()
    return blocks


def _append_field(paragraph, instr, sample="1"):
    """往段落末尾塞一个 w:fldSimple 域（页码这类只能由 Word 算的东西）。"""
    field = OxmlElement("w:fldSimple")
    field.set(qn("w:instr"), instr)
    run = OxmlElement("w:r")
    rpr = OxmlElement("w:rPr")
    size = OxmlElement("w:sz")
    size.set(qn("w:val"), "18")  # 9pt，与正文页脚一致
    rpr.append(size)
    run.append(rpr)
    text = OxmlElement("w:t")
    text.text = sample
    run.append(text)
    field.append(run)
    paragraph._p.append(field)
    return field


def _own_footer(section, total=False, values=None):
    """给答题卡/答案页单独一份页脚：本节从 1 开始编。

    实测老师的吐槽：8 开卷子 6 张，页脚写「共 8 页」（把 A4 答题卡也算进去了）；
    翻到答题卡又是「第 7 页」。所以这几节重新编号，并且**不印总页数**：
    SECTIONPAGES（本节页数）在 Word 里对，但 LibreOffice 转 PDF 时恒等于 1，
    老师拿到的 PDF 就成「第 2 页 共 1 页」—— 宁可不印，也不印错的。
    """
    section.footer.is_linked_to_previous = False
    footer = section.footer
    paragraph = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for child in list(paragraph._p):
        if child.tag != qn("w:pPr"):
            paragraph._p.remove(child)

    def text(value):
        run = paragraph.add_run(value)
        run.font.size = Pt(9)

    text("第 ")
    _append_field(paragraph, " PAGE ")
    text(" 页")
    if total:
        text(" 共 ")
        _append_field(paragraph, " SECTIONPAGES ")
        text(" 页")
    values = values or {}
    class_name = str(values.get("className") or values.get("class") or "").strip() or "__________"
    teacher = str(values.get("teacherName") or values.get("teacher") or "").strip() or "__________"
    text("　　使用班级：" + class_name + "　　出卷老师：" + teacher)
    # 本节开了"首页不同"（卷头在第一页页眉）时，第 1 页用的是 first_page_footer：
    # 只写 section.footer 的话第 1 页一个页脚都没有（实测正文/答题卡/答案页的第 1 页全丢）。
    try:
        different_first = bool(section.different_first_page_header_footer)
    except Exception:
        different_first = False
    if different_first:
        target = section.first_page_footer
        try:
            target.is_linked_to_previous = False
        except Exception:
            pass
        for paragraph_ in list(target.paragraphs):
            element = paragraph_._element
            if element.getparent() is not None:
                element.getparent().remove(element)
        for paragraph_ in footer.paragraphs:
            target._element.append(copy_module.deepcopy(paragraph_._element))


def _strip_page_total(document):
    """多节卷子里把模板页脚的「共 N 页」摘掉 —— 数字是错的，宁可不印。

    LibreOffice 算 NUMPAGES 时把「答题卡/答案页各自从 1 重新编号」也算进去：
    实测 4 页的卷子，页脚印「共 5 页」。Word 里是对的、PDF 里是错的，
    老师拿到手上对不上就更糟（与正文自带页脚「不印总页数」同一口径）。
    只动 footer / first_page_footer / even_page_footer 里的 fldSimple NUMPAGES。
    """
    for section in document.sections:
        for name in ("footer", "first_page_footer", "even_page_footer"):
            part = getattr(section, name, None)
            if part is None:
                continue
            for paragraph in part.paragraphs:
                for field in list(paragraph._p.findall(qn("w:fldSimple"))):
                    if "NUMPAGES" not in (field.get(qn("w:instr")) or "").upper():
                        continue
                    previous = field.getprevious()
                    following = field.getnext()
                    if previous is not None and previous.tag == qn("w:r"):
                        _trim_leading_char(previous, "共", from_end=True)
                    if following is not None and following.tag == qn("w:r"):
                        _trim_leading_char(following, "页", from_end=False)
                    paragraph._p.remove(field)


def _trim_leading_char(element, char, from_end):
    """把 run 末尾（或开头）的那个字去掉，留着前面的空格（「页 共」→「页 」）。"""
    nodes = element.findall(qn("w:t"))
    if not nodes:
        return
    node = nodes[-1] if from_end else nodes[0]
    text = node.text or ""
    if from_end:
        stripped = text.rstrip()
        if stripped.endswith(char):
            node.text = stripped[: -len(char)] + text[len(stripped):]
    else:
        stripped = text.lstrip()
        if stripped.startswith(char):
            node.text = text[: len(text) - len(stripped)] + stripped[len(char):]


def _restart_page_numbering(section, start=1):
    """本节页码从 start 重新开始（w:pgNumType w:start）。"""
    sect_pr = section._sectPr
    pg = sect_pr.find(qn("w:pgNumType"))
    if pg is None:
        pg = sect_pr.makeelement(qn("w:pgNumType"), {})
        # CT_SectPr 的元素顺序有要求：pgNumType 必须在 cols 之前，否则 Word 报文件损坏
        cols = sect_pr.find(qn("w:cols"))
        if cols is not None:
            cols.addprevious(pg)
        else:
            sect_pr.append(pg)
    pg.set(qn("w:start"), str(max(1, int(start))))


RE_ESSAY = re.compile(r"写作|作文|书面表达|写作题|作文题")


def _is_essay_section(section):
    """作文/书面表达这种「要学生写一大段」的大题：答题卡上得给作答格子，不能只给题号。"""
    if any(question.get("options") for question in section["questions"]):
        return False
    title = (section.get("label") or "") + (section["title"] or "")
    return bool(RE_ESSAY.search(title))


def _writing_area(document, section, lines=12, line_height_mm=9):
    """答题卡上的作答格：一格一行的带框表格，行高按手写留够。"""
    rows = max(4, int(lines))
    table = document.add_table(rows=rows, cols=1)
    table.style = "Table Grid"
    for row in table.rows:
        row.height = Mm(line_height_mm)
        row.height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST
    fit_table(table, column_width_twips(section))
    return table


RE_FILLIN = re.compile(r"填空|简答|完成句子|按首字母|首字母|句型转换|翻译|词语运用|语法填空|任务型|补全对话|用所给词|适当形式|根据汉语")

# 答案是一个字母的大题：完形、七选五、信息匹配，答题卡只该给题号格，不该给两行书写区
RE_LETTER_SECTION = re.compile(r"完形|cloze|七选五|信息匹配|匹配题|选字母", re.I)


def _is_fill_section(section):
    """填空/简答这类「答案是一小段字」的大题：答题卡要给带编号的作答格，只给题号格学生没处写。"""
    if _is_essay_section(section):
        return False
    title = (section.get("label") or "") + (section["title"] or "")
    if RE_LETTER_SECTION.search(title):
        return False
    if RE_FILLIN.search(title):
        return True
    # 标题没写明、但整节都没有选项 → 当填空处理（选择题一定有选项）
    return all(not question.get("options") for question in section["questions"])


def _filling_area(document, section, questions, start=1, lines=2, line_height_mm=8):
    """填空/简答题的作答格：左边题号、右边一条一块的写答案区。

    每块高度按手写留够（默认 2 行 × 8mm），学生写「He is tall.」不会挤到下一题格里。
    """
    rows = max(1, len(questions))
    table = document.add_table(rows=rows, cols=2)
    table.style = "Table Grid"
    for index, row in enumerate(table.rows):
        row.height = Mm(line_height_mm * max(1, int(lines)))
        row.height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST
        row.cells[0].text = str(start + index)
    width = column_width_twips(section)
    fit_table(table, width)
    # 题号只占一条窄边，其余留给学生写字：光改 cell.width 不管用（Word 按 tblGrid 排），
    # 要同时关掉自动伸缩并写列宽
    narrow = max(500, int(width * 0.07))
    table.autofit = False
    table.allow_autofit = False
    table.columns[0].width = Twips(narrow)
    table.columns[1].width = Twips(max(1000, width - narrow))
    for row in table.rows:
        row.cells[0].width = Twips(narrow)
        row.cells[1].width = Twips(max(1000, width - narrow))
    return table

def _set_section_page(section, paper="A4", orientation="portrait", margins=(20, 20, 18, 18), columns=1):
    width, height = PAPERS[paper]
    if orientation == "landscape":
        width, height = height, width
    section.page_width = Mm(width)
    section.page_height = Mm(height)
    section.top_margin = Mm(margins[0])
    section.bottom_margin = Mm(margins[1])
    section.left_margin = Mm(margins[2])
    section.right_margin = Mm(margins[3])
    # add_section 会克隆上一节的 sectPr：答题卡是 A4 竖版单栏，不加这几行就会继承卷子的
    # landscape 和两栏，答题表被压进半栏里
    sect_pr = section._sectPr
    cols = sect_pr.find(qn("w:cols"))
    if cols is None:
        cols = sect_pr.makeelement(qn("w:cols"), {})
        sect_pr.append(cols)
    cols.set(qn("w:num"), str(max(1, int(columns))))
    if max(1, int(columns)) == 1:
        cols.attrib.pop(qn("w:space"), None)
    pgsz = sect_pr.find(qn("w:pgSz"))
    if pgsz is not None:
        if orientation == "landscape":
            pgsz.set(qn("w:orient"), "landscape")
        else:
            pgsz.attrib.pop(qn("w:orient"), None)


def _score_value(value, fallback=2.0):
    """每题分值：写「2」「2分」都认，写错字要说清楚，别让它冒英文 traceback。"""
    if value in (None, ""):
        return fallback
    match = re.match(r"^\s*(\d+(?:\.\d+)?)\s*分?\s*$", str(value).strip())
    if not match:
        raise EngineError(
            "bad-score",
            "每题分值要填数字：「" + str(value) + "」不算",
            "写 2 或 2.5 都行；不想每题给分就先在大题标题里写「每题 2 分」")
    return float(match.group(1))


def build_exam(payload):
    """payload = {text, outputPath, title?, page?, baseTemplate?, values?, options?}

    也接受 payload["_parsed"]：A/B 卷两边要共用同一份解析结果，不能再走一遍
    「解析 → 写成文字 → 再解析」（那一趟会把完形短文、听力指令这些 notes 全丢掉）。
    """
    text = payload.get("text")
    options = dict(payload.get("options") or {})
    # 面板有时把 minutes / answerCard 直接放在顶层：两种写法都认，免得老师填了时长却没印出来
    for key in ("minutes", "answerCard", "answerSheet", "scorePerQuestion", "headerLine",
                "inlineAnswers", "answerDetail", "essayLines", "cardPaper", "cardMarginsMm",
                "answerPaper", "answerMarginsMm"):
        if key in payload and key not in options:
            options[key] = payload[key]
    parsed = payload.get("_parsed")
    if not (isinstance(parsed, dict) and isinstance(parsed.get("sections"), list)):
        parsed = parse_paper(text)
    warnings = list(parsed.get("warnings") or [])
    output = payload.get("outputPath") or default_output(
        "出卷", payload.get("title") or options.get("title") or "测试卷", ".docx")

    default_score = _score_value(options.get("scorePerQuestion"))
    total = parsed["totalScore"]
    for section in parsed["sections"]:
        for question in section["questions"]:
            if question["score"] is None:
                question["score"] = default_score
    total = round(
        sum(question["score"] or 0 for section in parsed["sections"] for question in section["questions"])
        + sum(float(section.get("declaredScore") or 0)
              for section in parsed["sections"] if section.get("writing")),
        2)

    if not any(section["questions"] or section.get("writing") for section in parsed["sections"]):
        # 一段没写题号的文字也能"出卷成功"，印出来只有标题和答题卡：老师得翻到最后一页才发现
        raise EngineError(
            "no-questions",
            "这段文字里没认出题目，先不出卷：" + (text or "")[:24].replace("\n", " ") + "…",
            "每道题前面要有题号（1. / 2、/（3）都认），大题写「一、选择题（每题 2 分）」；"
            "从别处粘题时注意题号别是图片或自动编号")

    # 卷名只认老师给的：以前缺省退成"第一个大题的标题"，卷首大字印成「一、单项选择」
    values = dict(payload.get("values") or {})
    title = payload.get("title") or options.get("title") or values.get("examTitle") or "测试卷"
    values.setdefault("examTitle", title)
    values.setdefault("totalScore", _fmt_score(total))
    minutes_value = options.get("minutes") or parsed.get("minutes")
    if minutes_value:
        values.setdefault("minutes", str(minutes_value))

    template_path = payload.get("baseTemplate")
    page = payload.get("page")
    if page is None and template_path:
        # 套了学校模板又没指定纸型：按模板自己的纸型/栏数走，别拿默认 8 开硬盖
        try:
            from .template_model import analyse_docx
            facts = analyse_docx(template_path).get("paper") or {}
            if facts.get("size"):
                page = {"size": facts["size"],
                        "orientation": facts.get("orientation"),
                        "columns": facts.get("columns") or 1}
        except Exception:
            page = None
    page = page or {"size": "8K", "orientation": "landscape", "columns": 2}
    paper_layout = resolve_paper(page)

    if template_path and payload.get("page"):
        # 既套了模板又手选了纸型：两边不一致时要说一声。老师以为"套了模板就是学校那份"，
        # 实际成品按她这次选的纸型排，容易印出来才发现不对。
        try:
            from .template_model import analyse_docx
            facts = analyse_docx(template_path).get("paper") or {}
            wide = abs(float(facts.get("widthMm") or 0) - float(paper_layout.get("widthMm") or 0))
            tall = abs(float(facts.get("heightMm") or 0) - float(paper_layout.get("heightMm") or 0))
            columns = int(facts.get("columns") or 1)
            if wide > 2 or tall > 2 or (columns > 1) != (int(paper_layout.get("columns") or 1) > 1):
                warnings.append({
                    "code": "template-paper-mismatch",
                    "message": "这份模板本身是 " + str(facts.get("size") or "另一种纸型")
                               + ("、" + str(columns) + " 栏" if columns > 1 else "、一栏")
                               + "，和你这次选的纸型不一样；成品按这次选的排，模板里的页眉页脚和字号照搬。"
                                 "想完全照模板就清空「纸型 / 版式」再点一次",
                })
        except Exception:
            pass

    title_blocks = [{"type": "paragraph", "align": "center", "runs": [{"text": title, "bold": True, "sizePt": 16}]}]
    caption = "（满分 " + _fmt_score(total) + " 分"
    if values.get("minutes"):
        caption += "，时间 " + str(values["minutes"]) + " 分钟"
    caption += "）"
    title_blocks.append({"type": "paragraph", "align": "center", "runs": [{"text": caption, "sizePt": 10.5}]})
    if options.get("headerLine", True):
        title_blocks.append({"type": "paragraph", "align": "center",
                            "runs": [{"text": "班级：__________  姓名：__________  学号：__________", "sizePt": 10.5}]})
    # 两栏卷子的卷头（卷名/分值/班级行）要跨整页宽。放正文里就得靠"连续分节"切栏数，
    # 而 LibreOffice 遇到正文中途改栏数的分节会把正文推下一页（长题干实测多印一张 8 开纸）。
    # 所以两栏时卷头写进「第一页页眉」：页眉天然跨整页，正文从下面直接两栏。
    header_blocks = []
    if paper_layout.get("columns", 1) > 1:
        header_blocks, blocks = title_blocks, []
    else:
        blocks = list(title_blocks)

    # 答题卡/答案页的勾选在排正文时就要知道：作文题要不要在卷面上留作答横线取决于它
    answer_card = options.get("answerCard") or "none"
    for section in parsed["sections"]:
        if not section["questions"]:
            if section.get("writing"):
                # 「二、书面表达」只有要求没有小题：标题、要求照印，再给老师一个明确的交代
                writing_heading = (section.get("label") or "") + (section["title"] or "")
                if section.get("declaredScore"):
                    writing_heading += "（" + _fmt_score(section["declaredScore"]) + " 分）"
                blocks.append({"type": "paragraph",
                               "runs": [{"text": writing_heading, "bold": True}], "keepWithNext": True})
                for line in section["notes"]:
                    blocks.append({"type": "paragraph", "runs": [{"text": line}]})
                if answer_card == "none":
                    blocks.append({
                        "type": "table",
                        "rows": [[""]] * max(6, int(options.get("essayLines") or 10)),
                        "rowHeightMm": 9,
                    })
                else:
                    blocks.append({"type": "paragraph",
                                   "runs": [{"text": "（请写在答题卡上）", "sizePt": 10.5}]})
                continue
            for line in section["notes"]:
                blocks.append({"type": "paragraph", "runs": [{"text": line}]})
            continue
        heading = (section.get("label") or "") + (section["title"] or "试题")
        section_score = sum(question["score"] or 0 for question in section["questions"])
        blocks.append({"type": "paragraph", "runs": [{"text": heading + "（" + _fmt_score(section_score) + " 分）", "bold": True}],
                       "keepWithNext": True})
        for line in section["notes"]:
            blocks.append({"type": "paragraph", "runs": [{"text": line}]})
        materials = section.get("materials") or []
        for index, question in enumerate(section["questions"]):
            for material in materials:
                if material.get("after") == index:
                    for line in material.get("lines") or []:
                        blocks.append({"type": "paragraph", "runs": [{"text": line}]})
            blocks.extend(_question_blocks(question, default_score))
            # 教师卷：答案与解析就印在题后（老师讲评时不翻页找答案）
            if options.get("inlineAnswers") and question.get("answer"):
                # 选项最后一行和【答案】拴在一起：否则答案会单独被挤到右栏顶上，
                # 老师看到的是"光秃秃一个【答案】B"，对不上是哪道题
                if blocks:
                    blocks[-1]["keepWithNext"] = True
                blocks.append({
                    "type": "paragraph",
                    "runs": [{"text": "【答案】" + str(question["answer"]), "bold": True, "sizePt": 10.5}],
                    "keepWithNext": bool(question.get("analysis")),
                })
                if question.get("analysis"):
                    blocks.append({
                        "type": "paragraph",
                        "runs": [{"text": "【解析】" + str(question["analysis"]), "sizePt": 10}],
                        "keepTogether": True,
                    })

    # 默认**不出**答题卡和答案页：老师没勾就不该多印两页（早先默认 card+end，
    # 一份 5 页的卷子点一下变成 10 页，和面板上的勾选框对不上）
    answer_sheet = options.get("answerSheet") or "none"
    has_any_answer = any(
        question.get("answer") for section in parsed["sections"] for question in section["questions"]
    )
    if answer_sheet != "none" and not has_any_answer:
        # 勾了答案页却一道答案都没解析到：**卷子照出**，只把答案页省掉并提醒。
        # （早先直接抛错拒绝出卷：老师 3 秒出不来卷子，比少一页答案页严重得多。）
        answer_sheet = "none"
        warnings.append({
            "code": "answers-missing",
            "message": "你勾了「附参考答案页」，但没找到任何答案，答案页就没印；"
                       "答案跟在每道题下面写「【答案】A」，或在卷尾写一段：1—5 BCBDA 6—10 ABABB",
        })

    # 正文 + 答题卡/答案页 = 一份文件里混了两种纸型。这时正文页脚不能印「共 N 页」：
    # 两种纸型的页数加在一起对谁都不对（8 开 6 张会被写成「共 8 页」）。
    extra_sections = answer_card != "none" or answer_sheet != "none"
    spec = {
        "outputPath": output,
        # 套了学校模板又没手选纸型：不要把上面探测出来的纸型再写回 spec。
        # 写回去会让 docx_build 按 resolve_paper 的**默认页边距**重排（模板 20mm → 24/27mm），
        # 老师明明套了学校模板，成品边距却不是学校的。page=None + 模板 = 原样保留模板版式。
        "page": page if (payload.get("page") or not template_path) else None,
        "blocks": blocks,
        "values": values,
        "overwrite": bool(payload.get("overwrite", False)),
        "defaultFooter": not extra_sections,
    }
    if header_blocks:
        spec["firstPageHeaderBlocks"] = header_blocks
    if payload.get("baseTemplate"):
        spec["baseTemplate"] = payload["baseTemplate"]
    if payload.get("footer"):
        spec["footer"] = str(payload["footer"])
    built = build_docx(spec)
    values = built.get("values") or values

    if extra_sections:
        document = docx.Document(output)
        if not payload.get("baseTemplate"):
            # 正文段页脚自己写（只有当前页，不印总页数）
            _own_footer(document.sections[0], total=False, values=values)
        if answer_card != "none":
            section = document.add_section(WD_SECTION.NEW_PAGE)
            # 新分节会从上一节抄 sectPr，连「首页不同」也抄过去：不关掉，答题卡第 1 页会顶着
            # 正文的卷头（标题+满分+班级行）再印一遍，整块卷头重复。
            section.different_first_page_header_footer = False
            _set_section_page(section, paper=options.get("cardPaper", "A4"), orientation="portrait",
                              margins=options.get("cardMarginsMm") or (20, 20, 18, 18))
            _own_footer(section, values=values)
            _restart_page_numbering(section)
            document.add_paragraph().add_run("答题卡").bold = True
            # 答题卡要跟卷子分开收：不印姓名栏，96 本收上来就分不清谁是谁
            document.add_paragraph().add_run(
                "班级：__________　姓名：__________　学号：__________　得分：__________")
            card_index = 0
            for block in parsed["sections"]:
                if not block["questions"]:
                    if block.get("writing"):
                        # 没有题号的作文大题：答题卡上也得有作答格，否则学生没处写
                        writing_head = document.add_paragraph()
                        writing_head.paragraph_format.keep_with_next = True
                        # 分页标在标题上（标在要求那段会把「二、书面表达」孤零零留在上一页）
                        if card_index:
                            writing_head.paragraph_format.page_break_before = True
                        writing_head.add_run((block.get("label") or "") + (block["title"] or ""))
                        notes = [line for line in block.get("notes") or [] if line.strip()]
                        if notes:
                            prompt = document.add_paragraph()
                            prompt.paragraph_format.keep_with_next = True
                            for line in notes[:-1]:
                                prompt.add_run(line)
                                prompt.add_run().add_break()
                            prompt.add_run(notes[-1])
                        _writing_area(document, section, lines=options.get("essayLines") or 10)
                    continue
                # 答题卡上不印分值：分值只在正文印一次
                card_heading = document.add_paragraph()
                card_heading.paragraph_format.keep_with_next = True  # 标题别孤零零留在上一页
                card_heading.add_run((block.get("label") or "") + (block["title"] or "试题"))
                if _is_essay_section(block):
                    # 作文题只印作答格（题号格留着也没用，学生不往里写字母）
                    first_number = card_index + 1
                    prompt = document.add_paragraph()
                    prompt.paragraph_format.keep_with_next = True  # 提示语别跟格子分到两页
                    # 作答格一定要整块落在一页上：跟着前面选择题的剩余空间走，12 行会被
                    # 拆成「这页 5 行 + 下页 7 行」，第二页九成空白（96 人每次多印一页）。
                    # 分页标在**标题**上：标在提示语上会把标题孤零零留在上一页页脚上方。
                    if card_index:
                        card_heading.paragraph_format.page_break_before = True
                    card_heading.paragraph_format.keep_with_next = True
                    prompt.add_run("第 %d 题　请在下面的格子里作答" % first_number)
                    _writing_area(document, section, lines=options.get("essayLines") or 10)
                    card_index += len(block["questions"])
                    continue
                if _is_fill_section(block):
                    # 填空/简答：一条一块的作答格（只给题号格学生没处写字）
                    _filling_area(document, section, block["questions"], start=card_index + 1,
                                  lines=options.get("fillLines") or 2)
                    card_index += len(block["questions"])
                    document.add_paragraph()
                    continue
                card_blocks = _card_blocks(block["questions"], start=card_index + 1)
                card_index += len(block["questions"])
                for rows in card_blocks:
                    table = document.add_table(rows=len(rows), cols=len(rows[0]))
                    table.style = "Table Grid"
                    for r_index, row in enumerate(rows):
                        for c_index, value in enumerate(row):
                            table.cell(r_index, c_index).text = value
                    fit_table(table, column_width_twips(section))
                if card_blocks:
                    document.add_paragraph()
        if answer_sheet != "none":
            # 答案页默认给老师看得懂的样子（题号 + 答案 + 题干摘要 + 解析）；
            # 只想要一行字母的老师可以传 answerDetail: False
            blocks_answer = _answer_blocks(parsed, detail=bool(options.get("answerDetail", True)))
            if blocks_answer:
                section = document.add_section(WD_SECTION.NEW_PAGE)
                section.different_first_page_header_footer = False
                _set_section_page(section, paper=options.get("answerPaper") or options.get("cardPaper", "A4"),
                                  orientation="portrait",
                                  margins=options.get("answerMarginsMm") or (20, 20, 18, 18))
                _own_footer(section, values=values)
                _restart_page_numbering(section)
                document.add_paragraph().add_run("参考答案 · " + str(title)).bold = True
                for line in blocks_answer:
                    if line.get("type") != "paragraph":
                        continue
                    paragraph = document.add_paragraph()
                    for run_spec in line.get("runs") or [{"text": line.get("text", "")}]:
                        run = paragraph.add_run(run_spec.get("text", ""))
                        if run_spec.get("bold"):
                            run.bold = True
        if template_path and extra_sections:
            # 模板页脚带「共 {{pages}} 页」：多节卷子里 LibreOffice 会算成错的数字，摘掉
            _strip_page_total(document)
        document.save(output)

    checked = preflight({"path": output, "values": values})
    question_count = sum(len(section["questions"]) for section in parsed["sections"])
    return {
        "outputPath": os.path.abspath(output),
        "questionCount": question_count,
        "sections": [
            {"title": section["title"], "label": section.get("label"),
             "questions": len(section["questions"]),
             "score": round(sum(question["score"] or 0 for question in section["questions"]), 2)
                      or round(float(section.get("declaredScore") or 0), 2),
             "writing": bool(section.get("writing")),
             "declaredScore": section.get("declaredScore")}
            for section in parsed["sections"] if section["questions"] or section.get("writing")
        ],
        "totalScore": total,
        "values": values,
        "warnings": list(warnings) + list(built.get("warnings") or []),
        "preflight": checked,
    }


def post_parse_warnings(parsed):
    return list(parsed.get("warnings") or [])

# ---------------------------------------------------------------- A/B 卷
# 真实需求：同一个班前后两场、或者同桌左右两列，老师要两份选项顺序不同的卷子，
# 防对答案。做法是把解析结果写成"粘贴格式"再走同一条出卷流水线，
# 保证 A/B 两套的排版、分值、答题卡完全同一套逻辑（别为 B 卷另写一遍排版）。

BIG_LABELS = ["一", "二", "三", "四", "五", "六", "七", "八", "九", "十",
              "十一", "十二", "十三", "十四", "十五"]


def serialize_paper(parsed, title=None):
    """把 parse_paper 的结果写回老师粘贴的那种格式（A/B 卷共用）。"""
    lines = []
    if title:
        lines.append(title)
    index = 0
    for section in parsed["sections"]:
        if not section["questions"] and not section.get("writing"):
            continue
        label = BIG_LABELS[index] if index < len(BIG_LABELS) else str(index + 1)
        index += 1
        declared = sum(float(question["score"] or 0) for question in section["questions"]) \
            or float(section.get("declaredScore") or 0)
        name = section.get("label") or ""
        heading = label + "、" + ((section.get("title") or "试题").strip() or "试题")
        if declared:
            heading += "（" + _fmt_score(declared) + " 分）"
        lines.append(heading)
        if section.get("writing"):
            # 作文只有要求没有小题：B 卷也得把要求带上，不然 B 卷那一段是空的
            for line in section.get("notes") or []:
                lines.append(line)
        for question in section["questions"]:
            stem = (question["stem"] or "").strip().replace("\n", " ")
            score = float(question["score"] or 0)
            head = str(question["number"]) + ". " + stem
            if score:
                head = head.rstrip() + "（" + _fmt_score(score) + " 分）"
            lines.append(head)
            for option in question["options"]:
                lines.append(option["label"] + ". " + option["text"])
            if question.get("answer"):
                lines.append("【答案】" + str(question["answer"]))
            if question.get("analysis"):
                lines.append("【解析】" + str(question["analysis"]).replace("\n", " "))
        lines.append("")
    return "\n".join(lines).strip() + "\n"


def _shuffle_options(parsed, seed=7):
    """B 卷：每道选择题的选项顺序打乱，答案字母跟着改写（题干、分值一律不动）。"""
    import random
    rng = random.Random(seed)
    changed = 0
    for section in parsed["sections"]:
        for question in section["questions"]:
            options = question.get("options") or []
            if len(options) < 2:
                continue
            answer = str(question.get("answer") or "").strip().upper()
            labels = [option["label"].upper() for option in options]
            if answer not in labels:
                continue          # 答案不是选项字母（填空/简答）：选项保持原样
            order = list(range(len(options)))
            rng.shuffle(order)
            if order == sorted(order):        # 洗成原样就再换一次，免得 B 卷和 A 卷一样
                order = order[1:] + order[:1]
            new_options = []
            new_labels = {}
            for new_index, old_index in enumerate(order):
                old = options[old_index]
                new_label = labels[new_index]
                new_labels[labels[old_index]] = new_label
                new_options.append({"label": new_label, "text": old["text"]})
            question["options"] = new_options
            question["answer"] = new_labels.get(answer, answer)
            changed += 1
    return changed


def build_exam_pair(payload):
    """出一份 A 卷 + 一份 B 卷（B 卷选项顺序打乱、答案跟着改）。"""
    options = dict(payload.get("options") or {})
    text = payload.get("text")
    parsed = parse_paper(text)
    base_title = (payload.get("title") or options.get("title")
                  or (parsed["sections"][0]["title"] if parsed["sections"] else None)
                  or "测试卷")
    # A 卷路径先定下来（没给就按标题落默认目录），B 卷才能在同一目录、只差一个名字
    first_output = payload.get("outputPath") or default_output(
        "出卷", base_title + "（A 卷）", ".docx")
    variants = []
    teacher_paths = []
    # 老师既要 A/B 防对答案、又要讲解用的教师卷：四份一起出（A / B / A 教师 / B 教师）
    teacher_versions = bool(options.get("teacher") or options.get("teacherCopy"))
    warnings = list(parsed.get("warnings") or [])
    first_path = None
    for label, seed in (("A", 0), ("B", 7)):
        copy = copy_module.deepcopy(parsed)
        if label == "B":
            changed = _shuffle_options(copy, seed)
            if not changed:
                warnings.append({"code": "ab-no-shuffle",
                                 "message": "这份卷子没有带选项的题，B 卷和 A 卷内容一样"})
        variant_title = base_title + "（" + label + " 卷）"
        variant_payload = {
            # 直接把解析结果交给 build_exam：早先这里写成 serialize_paper(copy) 再解析一遍，
            # 而 serialize 不写 notes / 材料段，A/B 卷就比正卷少了 12 段（听力指令、
            # 完形短文、阅读 A/B 篇全没了），只剩题号和选项 —— 印出来就是废卷。
            "text": text,
            "_parsed": copy,
            "title": variant_title,
            "page": payload.get("page"),
            "values": dict(payload.get("values") or {}),
            "options": dict(options),
        }
        if payload.get("baseTemplate"):
            variant_payload["baseTemplate"] = payload["baseTemplate"]
        # 路径由调用方按 base_title 定，这里统一算出来，A/B 放同一目录、名字只差一个字母
        variant_payload["outputPath"] = first_output if label == "A" else _variant_path(first_output, label)
        result = build_exam(variant_payload)
        missing = _missing_material(result["outputPath"], copy)
        if missing:
            warnings.append({
                "code": "ab-materials-lost",
                "message": "这份卷子有 " + str(len(missing)) + " 段材料没印出来（" + missing[0][:20] + "…），请先别印",
            })
        entry = {
            "label": label,
            "title": variant_title,
            "outputPath": result["outputPath"],
            "questionCount": result["questionCount"],
            "totalScore": result["totalScore"],
            "preflight": result["preflight"],
        }
        if teacher_versions:
            entry["teacherPath"] = _teacher_copy(variant_payload, variant_title, result["outputPath"])
            teacher_paths.append(entry["teacherPath"])
        variants.append(entry)
        if label == "A":
            first_path = result
    return {
        "outputPath": first_path["outputPath"],
        "variants": variants,
        "questionCount": first_path["questionCount"],
        "sections": first_path["sections"],
        "totalScore": first_path["totalScore"],
        "values": first_path["values"],
        "warnings": list(first_path.get("warnings") or []) + warnings,
        "preflight": first_path["preflight"],
        "ab": True,
        "teacherPaths": teacher_paths,
    }


def _missing_material(path, parsed):
    """出卷后自检：节说明与阅读材料有没有真的印进 docx。

    A/B 卷丢材料那次，preflight 全绿、卷子只有题号和选项 —— 光看返回值发现不了。
    这里把解析出来的材料段逐条拿去 docx 正文里对，缺了就报警（文档里 8 个条件）。
    """
    try:
        document = docx.Document(path)
    except Exception:
        return []
    body = "\n".join([paragraph.text for paragraph in document.paragraphs])
    for table in document.tables:
        for row in table.rows:
            body += "\n" + " ".join(cell.text for cell in row.cells)
    compact = re.sub(r"\s+", "", body)
    missing = []
    for section in parsed.get("sections") or []:
        samples = list(section.get("notes") or [])
        for material in section.get("materials") or []:
            samples.extend(material.get("lines") or [])
        for line in samples:
            probe = re.sub(r"\s+", "", line or "")
            if len(probe) < 8:
                continue
            # 开头 24 字对得上、尾 16 字对不上 = 材料只印了一半（被裁断/被当成标题合并了）
            tail = probe[-16:] if len(probe) > 48 else ""
            if probe[:24] not in compact or (tail and tail not in compact):
                missing.append(line)
    return missing


def _variant_path(path, label):
    """A 卷路径 → B 卷路径（同一目录，文件名尾巴加（B 卷））。"""
    if path:
        base, extension = os.path.splitext(str(path))
        renamed = base.replace("（A 卷）", "（B 卷）").replace("(A 卷)", "（B 卷）")
        if renamed == base:
            # 文件名里没有「A 卷」字样（老师自己写的路径）：加个 -B 后缀，别变成 "A（B 卷）"
            renamed = base + "-B"
        return renamed + extension
    return default_output("出卷", "测试卷（" + label + " 卷）", ".docx")



def _teacher_path(path):
    root, extension = os.path.splitext(path)
    return root + "（教师卷）" + extension


def _teacher_copy(variant_payload, variant_title, student_path):
    """A/B 卷各自的教师卷：答案与解析印在题后，不带答题卡也不另起答案页。"""
    options = dict(variant_payload.get("options") or {})
    options["inlineAnswers"] = True
    options["answerCard"] = "none"
    options["answerSheet"] = "none"
    options.setdefault("headerLine", False)
    options.pop("teacher", None)
    options.pop("teacherCopy", None)
    teacher_payload = dict(variant_payload)
    teacher_payload["options"] = options
    teacher_payload["title"] = variant_title + "（教师卷）"
    teacher_payload["outputPath"] = _teacher_path(student_path)
    teacher = build_exam(teacher_payload)
    return teacher["outputPath"]


def build_exam_teacher(payload):
    """学生卷 + 教师卷一次出两份。

    学生卷还是原来那一份（可带答题卡与答案页）；教师卷把【答案】/【解析】直接印在每道题下面，
    不印答题卡也不另起答案页 —— 老师讲评时不用来回翻。
    """
    student = build_exam(dict(payload))
    title = (student.get("values") or {}).get("examTitle") or payload.get("title") or "测试卷"
    teacher_payload = dict(payload)
    options = dict(payload.get("options") or {})
    options["inlineAnswers"] = True
    options["answerCard"] = "none"
    options["answerSheet"] = "none"
    # 教师卷是自己讲评用的，不印「班级/姓名/学号」那行（老师要印才显式打开）
    options.setdefault("headerLine", False)
    teacher_payload["options"] = options
    teacher_payload["title"] = title + "（教师卷）"
    teacher_payload["outputPath"] = _teacher_path(payload.get("outputPath") or student["outputPath"])
    teacher = build_exam(teacher_payload)
    student["teacherPath"] = teacher["outputPath"]
    student["secondLabel"] = "Word（教师卷）"
    student["teacher"] = {
        "outputPath": teacher["outputPath"],
        "questionCount": teacher["questionCount"],
        "totalScore": teacher["totalScore"],
        "preflight": teacher["preflight"],
    }
    return student


def exam_parse(payload):
    """面板"粘贴题干 → 看结构"：先给老师看一眼解析结果，再决定要不要出卷。

    分值要和出卷时用同一套默认值（没标的题按 scorePerQuestion，缺省 2 分），
    否则面板显示 0 分、成品写 4 分，老师会以为工具在乱算。
    """
    parsed = parse_paper(payload.get("text"))
    options = payload.get("options") or {}
    default_score = options.get("scorePerQuestion")
    default_score = float(default_score) if default_score else 2.0
    preview = []
    question_count = 0
    total = 0.0
    for section in parsed["sections"]:
        questions = []
        section_total = 0.0
        for question in section["questions"]:
            score = question["score"] if question["score"] is not None else default_score
            section_total += score
            question_count += 1
            questions.append({
                "number": question["number"],
                "stem": question["stem"][:120],
                "options": len(question["options"]),
                "hasAnswer": bool(question.get("answer")),
                "score": score,
            })
        total += section_total
        preview.append({"title": section["title"], "label": section.get("label"),
                        "declaredScore": section["declaredScore"],
                        "score": round(section_total, 2),
                        "notes": section["notes"], "questions": questions})
    return {
        "sections": preview,
        "questionCount": question_count,
        "totalScore": round(total, 2),
        "declaredTotalScore": parsed.get("declaredTotalScore"),
        "warnings": list(post_parse_warnings(parsed)),
    }
r"""印前检查：把"发下去会被教研组长打回来"的坑提前找出来。

判据不是拍脑袋，来自事实调研报告
（D:\eepseek Harness Workspace\中学教师办公AI助手_事实调研报告.md）里逐条带来源的规格：
- 试卷常用正度 8 开 368×260mm / 16 开 260×184mm / A3 297×420mm，"8K 或 A3 横向分两栏"
- 页边距：左右 2.7cm、上下 2.4cm
- 页脚：第 X 页 共 Y 页 + 使用班级 + 出卷老师

这里只做能自动判、判错了老师会说"这不是问题"的检查。宁少勿滥：
拿不准的一律 info（说一句，不扣分），只有"纸型/版心/页码域/占位符没填"才算 error。
"""

import re

from docx import Document
from docx.oxml.ns import qn

from .docx_inspect import inspect_docx
from .paper import MARGINS, PAPERS
from .units import emu_to_mm

# 学校卷子上常见的字体。不在这个表里的不一定是错（老师可能装了别的字库），
# 但换台电脑打开容易变形，值得提一句。
SAFE_FONTS = {
    "宋体", "黑体", "楷体", "楷体_GB2312", "仿宋", "仿宋_GB2312", "微软雅黑",
    "等线", "新宋体", "华文中宋", "华文楷体",
    "Times New Roman", "Arial", "Calibri", "Cambria", "Cambria Math", "Symbol",
}

# 章节序号：中文数字的"一、二、"与阿拉伯数字的"1. 2."
RE_BIG = re.compile(r"^\s*([一二三四五六七八九十]+)\s*[、.．]")
# 「（一）文言文阅读」是卷内小节，不是大题：题号可以在小节里重新从 1 起
RE_BIG_PAREN = re.compile(r"^\s*[（(]\s*([一二三四五六七八九十]+)\s*[)）]")
RE_ARABIC = re.compile(r"^\s*(\d{1,3})\s*(?:[、)）]|\.(?!\d)|．(?!\d))\s*\S")
RE_SCORE = re.compile(r"[（(]\s*(\d+(?:\.\d+)?)\s*分\s*[)）]")
RE_TOTAL = re.compile(r"满分\s*(\d+(?:\.\d+)?)\s*分")
RE_TOKEN = re.compile(r"\{\{\s*([A-Za-z0-9_]+)\s*\}\}")
CN_DIGITS = {ch: i + 1 for i, ch in enumerate("一二三四五六七八九")}
CN_DIGITS["十"] = 10


def _cn_number(text):
    """一 → 1，十二 → 12，二十三 → 23。"""
    if text == "十":
        return 10
    if text.startswith("十"):
        return 10 + CN_DIGITS.get(text[1], 0)
    if "十" in text:
        head, _, tail = text.partition("十")
        return CN_DIGITS.get(head, 0) * 10 + (CN_DIGITS.get(tail, 0) if tail else 0)
    return CN_DIGITS.get(text, 0)


def _finding(findings, level, code, message, hint=None, where=None):
    item = {"level": level, "code": code, "message": message}
    if hint:
        item["hint"] = hint
    if where:
        item["where"] = where
    findings.append(item)


def _page_size_name(width_mm, height_mm, tolerance=2.0):
    """按宽高认纸型；认不出返回 None。横向/纵向都认。"""
    w, h = sorted((width_mm, height_mm))
    for name, (pw, ph) in PAPERS.items():
        if abs(w - min(pw, ph)) <= tolerance and abs(h - max(pw, ph)) <= tolerance:
            return name
    return None


def _field_instructions(element):
    """拼出完整域指令：instrText 常被 Word 按 run 拆散（"PA"+"GE"），必须按 fldChar 结构聚合。"""
    out, buf, inside = [], [], False
    for node in element.iter():
        tag = node.tag.rsplit("}", 1)[-1]
        if tag == "fldSimple":
            instr = (node.get(qn("w:instr")) or "").strip()
            if instr:
                out.append(instr)
        elif tag == "fldChar":
            kind = node.get(qn("w:fldCharType"))
            if kind == "begin":
                inside, buf = True, []
            elif kind == "separate":
                inside = False
            elif kind == "end":
                if buf:
                    out.append("".join(buf).strip())
                inside, buf = False, []
        elif tag == "instrText" and inside:
            buf.append(node.text or "")
    return out


def _has_field(instructions, name):
    """域身份按 token 精确匹配：NUMPAGES / SECTIONPAGES 不算 PAGE。"""
    target = name.upper()
    for instr in instructions:
        for part in re.split(r'[\s"]+', instr.upper()):
            if part == target:
                return True
    return False


def _container_text(container):
    """页眉/页脚的文字：段落 + 表格单元格（页脚常做成一行表格）。"""
    parts = []
    try:
        for paragraph in container.paragraphs:
            if paragraph.text.strip():
                parts.append(paragraph.text.strip())
        for table in container.tables:
            for row in table.rows:
                cells = [cell.text.strip() for cell in row.cells]
                if any(cells):
                    parts.append(" ".join(c for c in cells if c))
    except Exception:
        pass
    return "\n".join(parts) or None


def _sections_facts(document):
    facts = []
    for index, section in enumerate(document.sections):
        footer_fields = _field_instructions(section.footer._element)
        header_fields = _field_instructions(section.header._element)
        facts.append({
            "index": index,
            "orientation": "横向" if section.orientation == 1 else "纵向",
            "pageWidthMm": round(emu_to_mm(section.page_width), 1),
            "pageHeightMm": round(emu_to_mm(section.page_height), 1),
            "marginsMm": {
                "top": round(emu_to_mm(section.top_margin), 1),
                "bottom": round(emu_to_mm(section.bottom_margin), 1),
                "left": round(emu_to_mm(section.left_margin), 1),
                "right": round(emu_to_mm(section.right_margin), 1),
            },
            "columns": _column_count(section),
            "footerText": _container_text(section.footer),
            "headerText": _container_text(section.header),
            "footerFields": footer_fields,
            "headerFields": header_fields,
        })
    return facts


def _column_count(section):
    try:
        cols = section._sectPr.find(qn("w:cols"))
        if cols is None:
            return 1
        return int(cols.get(qn("w:num")) or 1)
    except Exception:
        return 1


def _body_texts(document):
    texts = []
    for paragraph in document.paragraphs:
        text = paragraph.text.strip()
        if text:
            texts.append(text)
    for table in document.tables:
        for row in table.rows:
            cells = [cell.text.strip() for cell in row.cells]
            if any(cells):
                texts.append(" ".join(cells))
    return texts


def _paper_columns(sections):
    """这张纸实际用了几栏：标题节是单栏、正文节才分栏，取同一张纸上最多的那栏。"""
    if not sections:
        return 1
    first = sections[0]
    same_paper = [
        s for s in sections
        if abs(s["pageWidthMm"] - first["pageWidthMm"]) <= 2
        and abs(s["pageHeightMm"] - first["pageHeightMm"]) <= 2
    ] or [first]
    return max(s["columns"] for s in same_paper)


def _check_paper(sections, findings, facts):
    first = sections[0]
    columns = _paper_columns(sections)
    name = _page_size_name(first["pageWidthMm"], first["pageHeightMm"])
    facts["paper"] = name or "自定义"
    if not name:
        _finding(
            findings, "warn", "paper-unknown",
            "纸张不是常见规格（%.0f×%.0fmm），打印机会按 A4 缩印"
            % (first["pageWidthMm"], first["pageHeightMm"]),
            "试卷一般用 8K（368×260）或 A3（420×297）横向",
        )
        return
    if name == "Letter":
        _finding(
            findings, "info", "paper-letter",
            "纸张是 Letter（216×279mm），国内打印机一般按 A4 走，右边会多一道白边",
            "出一份 A4 或 8K/A3 的版更稳",
        )
    wide = max(first["pageWidthMm"], first["pageHeightMm"]) > 300
    if wide and first["orientation"] != "横向":
        _finding(
            findings, "warn", "paper-orientation",
            "%s 用成了纵向，印出来会窄一半" % name,
            "试卷请改横向（宽 %.0fmm）" % max(first["pageWidthMm"], first["pageHeightMm"]),
        )
    if wide and columns < 2:
        _finding(
            findings, "warn", "paper-columns",
            "%s 横向只有 1 栏，一行字太宽老师看着累" % name,
            "试卷常规是两栏",
        )
    if columns >= 2 and not wide:
        _finding(
            findings, "info", "paper-columns-narrow",
            "%s 用了 %d 栏，栏宽只有 %.0fmm" % (name, columns, first["pageWidthMm"] / max(columns, 1)),
        )


# 下载站往页眉里塞的推广语。老师自己打出来的页眉不会写这些。
AD_MARKERS = ("免费下载", "请勿商用", "资源网", "课件下载", "下载地址", "www.", "http://", "https://")
# 学校抬头的特征词：带这些的页眉一律留（学校官网也是学校的）
SCHOOL_MARKERS = ("中学", "小学", "学校", "附中", "附小", "教研", "年级", "备课", "命题", "出卷", "学期", "考试")


def _check_header(sections, findings):
    """页眉里有没有下载站硬塞的广告（会跟着卷子一起打印、发给学生）。"""
    seen = set()
    for section in sections:
        text = (section.get("headerText") or "").strip()
        if not text or text in seen:
            continue
        seen.add(text)
        if any(marker in text for marker in SCHOOL_MARKERS):
            continue
        if not any(marker in text.lower() if marker.isascii() else marker in text for marker in AD_MARKERS):
            continue
        _finding(
            findings, "warn", "header-ad",
            "页眉里有下载站的推广语：" + text[:40],
            "文件加工里选「净化」会自动删掉；自己留着的学校抬头不会动",
        )


def _check_margins(sections, findings, facts):
    first = sections[0]
    name = facts.get("paper")
    expected = MARGINS.get(name)
    if not expected:
        return
    top, bottom, left, right = expected
    got = first["marginsMm"]
    facts["marginsExpectedMm"] = {"top": top, "bottom": bottom, "left": left, "right": right}
    deviations = []
    for key, want in (("left", left), ("right", right), ("top", top), ("bottom", bottom)):
        if abs(got[key] - want) > 1.0:
            deviations.append("%s %.1fmm（常规 %.1fmm）" % ({"left": "左", "right": "右", "top": "上", "bottom": "下"}[key], got[key], want))
    if deviations:
        _finding(
            findings, "warn", "margin-off",
            "页边距和常规不一样：" + "，".join(deviations),
            "试卷常规是左右 2.7cm、上下 2.4cm；要装订就留宽左边距",
        )


def _check_footer(sections, findings, facts):
    texts, fields = [], []
    for section in sections:
        # 分节后每节都引用同一份页脚，文字会重复两遍 —— 去重后再看
        if section["footerText"] and section["footerText"] not in texts:
            texts.append(section["footerText"])
        for q in section["footerFields"]:
            value = q.strip().upper()
            if value and value not in fields:
                fields.append(value)
    footer_text = "\n".join(texts)
    has_page = _has_field(fields, "PAGE")
    # 多纸型混排时答题卡/答案页用 SECTIONPAGES（本节页数）重新编号，
    # 有它就算「写了总页数」—— 别让老师为了过检去把本节页数改成全文页数
    has_pages = _has_field(fields, "NUMPAGES") or _has_field(fields, "SECTIONPAGES")
    facts["footer"] = {
        "text": footer_text or None,
        "fields": fields,
        "hasPageField": has_page,
        "hasTotalField": has_pages,
    }
    if not footer_text and not fields:
        _finding(
            findings, "warn", "footer-missing",
            "没有页脚，学生翻页掉了就对不上号",
            "页脚放：第 X 页 共 Y 页 · 使用班级 · 出卷老师",
        )
        return
    if not has_page:
        _finding(
            findings, "error", "footer-no-page-field",
            "页脚的页码是手打的字，不是页码域（改内容后页码不会跟着变）",
            "用「第 {PAGE} 页 共 {NUMPAGES} 页」这样的域；模板里 {{page}}/{{pages}} 会自动换成域",
        )
    elif not has_pages:
        _finding(
            findings, "info", "footer-no-total",
            "页脚只有当前页，没有总页数",
            "加上「共 {NUMPAGES} 页」，老师印完能核对张数",
        )
    if has_page and len(sections) > 1:
        lonely = [
            str(section["index"] + 1) for section in sections
            if section["footerText"] and not _has_field(section["footerFields"], "PAGE")
        ]
        if lonely:
            _finding(
                findings, "info", "footer-section-no-page",
                "第 %s 节页脚没有页码域" % "、".join(lonely),
                "多节文档（比如正文 + 答题卡）每节页脚都放一次页码域，不然翻到后面没页码",
            )
    if footer_text and not re.search(r"班", footer_text):
        _finding(findings, "info", "footer-no-class", "页脚没写使用班级", "多班用同一份卷子时容易拿错")
    if footer_text and not re.search(r"老师|教师|出卷|命题", footer_text):
        _finding(findings, "info", "footer-no-teacher", "页脚没写出卷老师", "学校归档常要求署名")


def _check_numbering(texts, findings):
    """小题编号的作用域是它所在的大题；大题允许回到 1（一个文件可以有两卷）。

    「（一）文言文阅读」这种卷内小节单独成组、不参与大题序列 —— 语文卷常有，
    把它当大题会一路报跳号，老师只能删自己的小标题才能过检。
    """
    sections, subdivisions, loose, current = [], [], [], None
    for text in texts:
        bare = RE_BIG.match(text)
        paren = RE_BIG_PAREN.match(text)
        if bare:
            current = {"big": _cn_number(bare.group(1)), "text": text[:24], "items": []}
            sections.append(current)
        elif paren:
            current = {"text": text[:24], "items": []}
            subdivisions.append(current)
        small = RE_ARABIC.match(text)
        if small:
            (current["items"] if current else loose).append((int(small.group(1)), text[:24]))

    _sequence([(s["big"], s["text"]) for s in sections], findings, "big", "大题", restart_ok=True)
    if loose:
        _sequence(loose, findings, "small", "小题", restart_ok=True)
    for section in sections:
        _sequence(section["items"], findings, "small", "小题", restart_ok=False)
    for section in subdivisions:
        _sequence(section["items"], findings, "small", "小题", restart_ok=True)


def _sequence(items, findings, kind, label, restart_ok):
    if not items:
        return
    values = [number for number, _ in items]
    # 小题在各大题内各自从 1 起是常态；只有整卷/大题序列从中间开始时才提一句。
    if values[0] != 1 and restart_ok:
        _finding(findings, "info", "numbering-start-" + kind, "%s从 %d 开始编号" % (label, values[0]))
    seen = set()
    previous = None
    for number, text in items:
        if number == 1 and previous is not None and restart_ok:
            seen, previous = {1}, 1
            continue
        if number in seen:
            _finding(findings, "error", "numbering-duplicate-" + kind, "%s编号重复：%d（%s）" % (label, number, text))
        seen.add(number)
        if previous is not None and number > previous + 1:
            missing = "、".join(str(n) for n in range(previous + 1, number))
            _finding(findings, "error", "numbering-gap-" + kind, "%s编号跳号：缺 %s（%s）" % (label, missing, text))
        previous = number


def _check_score(document, texts, values, findings, facts):
    inline_total = None
    small_total = small_count = 0.0
    big_total = big_count = 0
    for text in texts:
        # "（满分 100 分，时间 90 分钟）" 是封面行，不是某道题的分
        total_match = RE_TOTAL.search(text)
        if total_match:
            if inline_total is None:
                inline_total = float(total_match.group(1))
            continue
        # （一）小节的分不是大题分，不参与大题分值口径
        is_big = bool(RE_BIG.match(text)) and not RE_BIG_PAREN.match(text)
        for match in RE_SCORE.finditer(text):
            if is_big:
                big_total += float(match.group(1))
                big_count += 1
            else:
                small_total += float(match.group(1))
                small_count += 1
    # 分值有的卷子写在小题上，有的只写在大题标题上；两处都加就翻倍（100 分变 200 分）。
    # 小题分优先，只有小题没标分时才用大题分。
    if small_count:
        total, counted = small_total, small_count
    else:
        total, counted = big_total, big_count
    declared = None
    if values and values.get("totalScore") not in (None, ""):
        try:
            declared = float(values["totalScore"])
        except (TypeError, ValueError):
            declared = None
    if declared is None:
        declared = inline_total
    facts["score"] = {"sum": round(total, 1), "counted": counted, "declared": declared}
    if counted == 0:
        _finding(findings, "info", "score-none", "没找到「（X 分）」这样的分值标注", "大题后面写分值，老师改卷时好分配时间")
        return
    if declared is None:
        _finding(findings, "info", "score-no-total", "题上标了分值，但没写满分", "标题下写「满分 X 分」便于核对")
        return
    if abs(total - declared) > 0.01:
        _finding(
            findings, "warn", "score-mismatch",
            "各题分值加起来 %.1f 分，写的满分是 %.1f 分，差 %.1f 分" % (total, declared, total - declared),
            "老师最常在这里被打回来：漏标一个大题的分值",
        )


def _check_fonts(info, findings, facts):
    used = [item["font"] for item in info.get("fonts") or [] if item.get("font")]
    facts["fonts"] = used
    unsafe = [font for font in used if font not in SAFE_FONTS]
    if unsafe:
        _finding(
            findings, "warn", "font-unsafe",
            "用了可能没装的字体：" + "、".join(unsafe[:5]),
            "换台电脑打开会变形，建议换成宋体/楷体/黑体",
        )


CONTAINERS = (
    "header", "footer",
    "first_page_header", "first_page_footer",
    "even_page_header", "even_page_footer",
)


def _every_text_node(document):
    """正文、表格、页眉页脚、文本框、脚注尾注里的每一段文字。

    以前只走 document.paragraphs + 表格：模板里文本框的 {{school}} 和脚注里的
    占位符都看不见，preflight 全绿而学生拿到照样能看见 {{...}}。
    """
    parts = [document.element.body]
    for section in document.sections:
        for name in CONTAINERS:
            try:
                parts.append(getattr(section, name)._element)
            except Exception:
                continue
    try:
        for part in document.part.package.iter_parts():
            if str(part.partname).endswith(("footnotes.xml", "endnotes.xml")):
                parts.append(part.element)
    except Exception:
        pass
    for element in parts:
        for node in element.iter(qn("w:t")):
            yield node.text or ""


def _check_tokens(document, sections, findings):
    leftovers = set()
    texts = list(_every_text_node(document))
    for text in texts:
        for match in RE_TOKEN.finditer(text or ""):
            leftovers.add(match.group(1))
    if leftovers:
        _finding(
            findings, "error", "placeholders-left",
            "模板占位符没填：" + "、".join(sorted(leftovers)),
            "这些位置学生拿到会直接看到 {{...}}",
        )


def _check_pictures(document, findings, facts):
    images = 0
    try:
        images = len(document.inline_shapes)
    except Exception:
        images = 0
    facts["images"] = images
    if images:
        _finding(findings, "info", "pictures", "文档里有 %d 张图片，印前翻一遍别糊了" % images)


def _check_answer_table(document, findings, facts):
    """答题卡的题号行要连续。

    答题卡是**按 10 格分块**排的（一块一张小表），一节的题号会落到好几张表里；
    早先只看每张表的第 0 行、且只收一行，于是 41 题的卡被读成 36 个带缺口的号，
    每份卷子都误报「答题框编号不连续」。现在把所有「题号」行都收进来。
    「答案」行是给学生填的空格，不参与。
    """
    numbers = []
    for table in document.tables:
        for row in table.rows:
            cells = [cell.text.strip() for cell in row.cells]
            if not cells or cells[0] != "题号":
                continue
            for cell in cells[1:]:
                if cell.isdigit():
                    numbers.append(int(cell))
    if not numbers:
        return
    facts["answerTableNumbers"] = numbers
    expected = list(range(1, len(numbers) + 1))
    if numbers != expected:
        _finding(
            findings, "warn", "answer-table-numbering",
            "答题框编号不是 1 到 %d 连续（%s）" % (len(numbers), "、".join(str(n) for n in numbers[:12])),
            "学生涂卡时对不上题号",
        )


RE_ANSWERS_HEAD = re.compile(r"^(参考答案|答案与解析|答案|参考答案与解析|答案及解析)$")


def _paper_texts(document, sections):
    """主卷正文的段落：标题节 + 正文节（同一张纸上的都算），不含答题卡/答案页。

    分栏卷子的标题块在单独一节（单栏），所以不能再"只取第一节"；
    答题卡和答案页换了纸（A4）或者开头就是「参考答案」，都要排除 ——
    那里的题号（答题框 1 2 3）和分值重复，算进去会误报。
    """
    from docx.oxml.ns import qn

    if not sections:
        return []
    first = sections[0]
    paper_indexes = {
        s["index"] for s in sections
        if abs(s["pageWidthMm"] - first["pageWidthMm"]) <= 2
        and abs(s["pageHeightMm"] - first["pageHeightMm"]) <= 2
    } or {0}
    texts, section_index = [], 0
    for child in document.element.body:
        if child.tag not in (qn("w:p"), qn("w:tbl")):
            continue
        # 表格也要算：题出在表格里（教案式卷子、答题格）时，以前整张表被跳过，
        # 题号/分值统计就只剩零星几行，score-mismatch 会误报。
        text = "".join(node.text or "" for node in child.iter(qn("w:t"))).strip()
        ppr = child.find(qn("w:pPr")) if child.tag == qn("w:p") else None
        ends_section = ppr is not None and ppr.find(qn("w:sectPr")) is not None
        if text and section_index in paper_indexes:
            if RE_ANSWERS_HEAD.match(text):
                break
            texts.append(text)
        if ends_section:
            section_index += 1
    # body 末尾那个 sectPr 就是最后一节的收尾，没有对应段落，不需要再 +1
    return texts


MAX_PER_CODE = 3


def _cap_findings(findings):
    """同类问题最多留 3 条，剩下的折成一句（1000 段的卷子否则会吐几十 KB）。"""
    kept, counts, levels = [], {}, {}
    for item in findings:
        code = item["code"]
        levels[code] = item["level"]
        counts[code] = counts.get(code, 0) + 1
        if counts[code] <= MAX_PER_CODE:
            kept.append(item)
    for code, total in counts.items():
        if total > MAX_PER_CODE:
            kept.append({
                "level": levels[code],
                "code": code + "-more",
                "message": "同类问题还有 %d 处，先按上面几处改" % (total - MAX_PER_CODE),
            })
    return kept


def preflight(payload):
    """payload = {path, values?} → 印前检查结果（findings + 分数）。"""
    from .errors import EngineError, require

    path = payload.get("path")
    require(bool(path), "path-required", "没有给出文件路径", "点「浏览…」选一份 .docx（印前检查的对象就是你选的文件）")
    import os

    full = os.path.abspath(path)
    require(os.path.isfile(full), "file-not-found", "找不到这个文件：" + full, "确认文件还在原位置；从 U 盘/微信里拷到电脑上再选")

    try:
        document = Document(full)
    except Exception as exc:
        raise EngineError(
            "invalid-docx",
            "这个文件打不开，可能不是 .docx 或者已经损坏：" + os.path.basename(full),
            "用 Word / WPS 打开它，另存为 .docx 再拖进来",
        ) from exc
    info = inspect_docx(full, preview=0)
    sections = _sections_facts(document)
    require(len(sections) > 0, "no-section", "这个文档没有可用的页面设置")
    texts = _body_texts(document)
    # 编号与分值只看主卷正文：答题卡/答案页会在第二节之后，那里重复的题号与分值会误报
    paragraph_texts = _paper_texts(document, sections)
    values = payload.get("values") or {}

    findings = []
    facts = {}
    _check_paper(sections, findings, facts)
    _check_margins(sections, findings, facts)
    _check_footer(sections, findings, facts)
    _check_header(sections, findings)
    if not paragraph_texts and not document.tables:
        _finding(
            findings, "warn", "content-empty",
            "文档里没有找到正文内容",
            "确认选的是要打印的那份文件；扫出来的空白页也会这样",
        )
    _check_numbering(paragraph_texts, findings)
    _check_score(document, paragraph_texts, values, findings, facts)
    _check_fonts(info, findings, facts)
    _check_tokens(document, sections, findings)
    _check_pictures(document, findings, facts)
    _check_answer_table(document, findings, facts)

    errors = [f for f in findings if f["level"] == "error"]
    warns = [f for f in findings if f["level"] == "warn"]
    score = max(0, 100 - 15 * len(errors) - 5 * len(warns))
    level = "fail" if errors else ("warn" if warns else "pass")
    findings = _cap_findings(findings)
    summary = {
        "pass": "印前检查通过，可以导出",
        "warn": "%d 项建议改（不拦导出）" % len(warns),
        "fail": "%d 项必须改，改了再导出" % len(errors),
    }[level]
    return {
        "ok": not errors,
        "level": level,
        "score": score,
        "summary": summary,
        "findings": findings,
        "facts": facts,
    }
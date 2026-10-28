r"""题库导入：老师手里的 Excel / Word 表格 → 一份能排版的卷子。

为什么单独开一个模块：
老师真实手里的题库不是一段文字，而是「一列题干、四列选项、一列答案」的 Excel，
或者 Word 里随手画的一张表格。让他们把两百道题一条条粘进对话框是不现实的。
所以这一层只做一件事：把各种形状的表格读成同一种「行」，
再吐成 exam.parse_paper 已经认识的粘贴文本。

排版（分栏、答题卡、答案页、印前检查）一概留在 exam.py，本模块不重写第二套
「选项怎么排」的规则 —— 两份规则早晚会改歪一份，而歪的那份没人测得到。
"""

import csv
import io
import os
import re

from .errors import EngineError, require
from .paths import default_output
from .profile import cn_ordinal

# 表头模糊匹配用词。中文按「包含」认（选项A / 第4选项 都能命中），
# 英文按「整词或包含」认：老师从各种题库网站导出的表头并不统一。
NUMBER_WORDS = ("题号", "序号", "编号", "小题号", "no", "num", "number", "id")
TYPE_WORDS = ("题型", "类型", "类别", "分类", "type", "kind")
STEM_WORDS = ("题干", "题目", "试题", "问题", "内容", "question", "stem", "content")
ANSWER_WORDS = ("答案", "answer", "正确答案", "参考答案", "key")
ANALYSIS_WORDS = ("解析", "分析", "详解", "点评", "analysis", "explain", "solution")
SCORE_WORDS = ("分值", "分数", "score", "points", "每题分")
OPTION_WORDS = ("选项", "备选项", "option", "options", "选项内容")
# 阅读/完形/任务型阅读的材料（短文）单独一列：题库里有，出卷时排在题目前面。
# persona 抓到过"裸题"：阅读题只印题干，学生找不到短文。
MATERIAL_WORDS = ("材料", "短文", "语篇", "篇章", "阅读材料", "文章", "素材",
                  "passage", "material", "article", "text")
# 这几类题没有材料就是裸题，要提醒老师
MATERIAL_TYPES = ("阅读", "完形", "任务型", "补全对话", "七选五", "短文填空", "reading", "cloze")

# 「选项A」「A」「a 项」都算第 A 个选项列；只认单字母，免得把「答案」当成选项。
RE_OPTION_HEADER = re.compile(r"^(?:选项)?([a-h])(?:项|选项|栏)?$")
# 一行里塞了多个选项：「A. 甲 B. 乙 C. 丙」。
# 字母前不能是英文字母/数字，否则「A book」里的 A 也会被当成选项。
RE_OPTION_MARK = re.compile(r"(?<![A-Za-z0-9])([A-Ha-h])\s*[、.．)）:：]\s*")
# 题干末尾已经写了分值时，别再叠一个「（2 分）」——叠了两个分，印前检查会报分对不上。
RE_TRAILING_SCORE = re.compile(r"[（(]\s*\d+(?:\.\d+)?\s*分\s*[)）]\s*$")
# 没有题号列时，题干自己带的「3、」「3. 」也认；但「3.14」不能认成第 3 题。
RE_STEM_NUMBER = re.compile(r"^\s*(\d{1,3})\s*(?:[、)）]|\.(?!\d)|．(?!\d))\s*")
# 大题标题里的分值说明：「选择题（共 10 小题，每小题 2 分）」→「选择题」
RE_TYPE_SCORE = re.compile(r"[（(][^（）()]{0,40}?\d+(?:\.\d+)?\s*分[^（）()]{0,40}?[)）]")
RE_TYPE_PREFIX = re.compile(
    r"^\s*(?:第\s*[一二三四五六七八九十]{1,3}\s*部分\s*[、.．:：]?"
    r"|[（(]?[一二三四五六七八九十ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ]{1,3}[)）、.．:：])\s*")
RE_HEADER_JUNK = re.compile(r"[\s\u3000*＊#]")
RE_HEADER_TRIM = "()[]【】、.:：,，/\\-—_"

DEFAULT_FALLBACK_HINT = "没认出表头，按 题号/题干/A-D/答案/解析/分值 读"
SUPPORTED = (".xlsx", ".docx", ".csv", ".txt")


# ---------------------------------------------------------------- 小工具


def _warn(code, message, **extra):
    payload = {"code": code, "message": message}
    payload.update(extra)
    return payload


def _dedupe_warnings(warnings):
    """同一个毛病只报一次：Word 里五张表都没表头，不该刷五条一样的提示。"""
    seen = set()
    kept = []
    for item in warnings:
        key = (item.get("code"), item.get("message"))
        if key in seen:
            continue
        seen.add(key)
        kept.append(item)
    return kept


def _cell_text(value):
    """单元格 → 文本。Excel 的 2.0 要变成「2」，否则分值列全是 2.0、题号全是 1.0。"""
    if value is None:
        return ""
    if isinstance(value, bool):
        return "对" if value else "错"
    if isinstance(value, float):
        if abs(value - round(value)) < 1e-9:
            return str(int(round(value)))
        return ("%.4f" % value).rstrip("0").rstrip(".")
    if isinstance(value, int):
        return str(value)
    text = str(value)
    return text.replace("\r\n", "\n").replace("\r", "\n").strip()


def _norm_header(text):
    value = _cell_text(text).lower()
    value = value.replace("（", "(").replace("）", ")")
    value = RE_HEADER_JUNK.sub("", value)
    return value.strip(RE_HEADER_TRIM)


def _has(text, words):
    return any(word in text for word in words)


def _header_kind(cell):
    """表头单元格 → (角色, 选项字母)。认不出返回 (None, None)。

    判定顺序有讲究：先认「解析」「答案」，再认「题号」「题型」，
    最后才认「题干」——「答案解析」必须落在解析列，
    「题型」也不能被「题目」抢走。
    """
    text = _norm_header(cell)
    if not text:
        return None, None
    match = RE_OPTION_HEADER.match(text)
    if match:
        return "option", match.group(1).upper()
    if _has(text, ANALYSIS_WORDS):
        return "analysis", None
    if _has(text, ANSWER_WORDS):
        return "answer", None
    if _has(text, NUMBER_WORDS):
        return "number", None
    if _has(text, TYPE_WORDS):
        return "type", None
    if _has(text, MATERIAL_WORDS):
        return "material", None
    # 合并选项列要排在「题干」前面：「选项内容」里有「内容」，
    # 先认题干就会被抢成一列题干；先认「选项」才是老师的意思。
    if _has(text, OPTION_WORDS):
        return "options", None
    if _has(text, STEM_WORDS):
        return "stem", None
    if _has(text, SCORE_WORDS):
        return "score", None
    return None, None


def _empty_roles():
    return {"number": None, "type": None, "stem": None, "options": {},
            "merged": None, "answer": None, "analysis": None, "score": None,
            "material": None}


def _header_map(row):
    """认出表头就返回列映射，认不出返回 None（调用方走顺序兜底）。"""
    roles = _empty_roles()
    found = 0
    for index, cell in enumerate(row):
        kind, letter = _header_kind(cell)
        if not kind:
            continue
        found += 1
        if kind == "option":
            roles["options"].setdefault(letter, index)
        elif kind == "options":
            if roles["merged"] is None:
                roles["merged"] = index
        elif roles[kind] is None:
            roles[kind] = index
    if not found:
        return None
    # 只认出一列弱信号（题号 / 单个「A」）时，多半是数据行而不是表头：
    # 无表头的题库，第一行答案常常就写成「A」，会被误当成 A 选项列。
    if roles["stem"] is not None or roles["merged"] is not None:
        return roles
    if len(roles["options"]) >= 2 or found >= 3:
        return roles
    return None


def _options_look_merged(rows, column):
    """某一列的内容是不是「A. 甲 B. 乙 C. 丙」这种合并写法？"""
    sample = [row[column] for row in rows[:20] if column < len(row) and row[column]]
    if not sample:
        return False
    merged = sum(1 for text in sample if len(RE_OPTION_MARK.findall(text)) >= 2)
    return merged * 2 >= len(sample)


def _fallback_roles(width, rows):
    """认不出表头时的顺序兜底：[题号, 题干, A-D, 答案, 解析, 分值]。

    列数不够九列时按列数截断；但只要发现「选项那一列」写的是合并的 A./B./C.，
    就把后面几列整体前移一位认（题号/题干/选项/答案/解析/分值）——
    老师自己整理的表格很少正好九列。
    """
    layout = ["number", "stem", "A", "B", "C", "D", "answer", "analysis", "score"]
    if width >= 3 and _options_look_merged(rows, 2):
        layout = ["number", "stem", "merged", "answer", "analysis", "score"]
    if width <= 1:
        layout = ["stem"]
    roles = _empty_roles()
    for index, name in enumerate(layout[:width]):
        if name in ("A", "B", "C", "D"):
            roles["options"][name] = index
        elif name == "merged":
            roles["merged"] = index
        else:
            roles[name] = index
    return roles


RE_ANSWER_CELL = re.compile(r"^(?:[A-Ha-h]{1,4}|对|错|正确|错误|√|×|T|F)$")


def _column_values(rows_in, index):
    return [row[index].strip() for row in rows_in[:40]
            if index < len(row) and row[index] and row[index].strip()]


def _answer_like(values):
    if not values:
        return False
    hits = sum(1 for value in values if RE_ANSWER_CELL.match(value))
    return hits * 10 >= len(values) * 7


def _small_number_like(values):
    if not values:
        return False
    for value in values:
        match = re.fullmatch(r"(\d+(?:\.\d+)?)\s*分?", value)
        if not match or not (0 < float(match.group(1)) <= 100):
            return False
    return True


def _average_length(values):
    return (sum(len(value) for value in values) / len(values)) if values else 0.0


def _repair_roles(roles, rows_in, width):
    """没表头时的列体检：按内容投票改映射，比死按列序切猜得准。

    老师自己整理的表格常是「题号 | 题干 | 答案 | 解析 | 分值」五列，顺序兜底会把
    答案当选项 A、解析当选项 B、分值当选项 C —— 题面直接印成「A. A B. 这是解析 C. 2」。
    整列都是「A/B/C/对/错」的那一列就是答案列：从这个位置往右重新认（分值、解析），
    往左把题干找回来。认不出来就原样返回，不做猜测。
    """
    if roles.get("stem") is None and not roles.get("options"):
        return roles
    answer_index = roles.get("answer")
    answer_like_options = any(_answer_like(_column_values(rows_in, index))
                              for index in roles["options"].values())
    if answer_index is None or answer_like_options:
        found = None
        for column in range(width - 1, 0, -1):
            if _answer_like(_column_values(rows_in, column)):
                found = column
                break
        if found is None:
            return roles
        answer_index = found
    if not _answer_like(_column_values(rows_in, answer_index)):
        return roles

    # 答案列右边的一律重认：不沿用兜底布局给的「B/C/D 选项」
    for name, index in list(roles["options"].items()):
        if index >= answer_index:
            del roles["options"][name]
    for name in ("analysis", "score", "answer"):
        if roles.get(name) is not None and roles[name] >= answer_index:
            roles[name] = None
    roles["answer"] = answer_index

    # 题干被挤到答案列右边时（如「题干 | 答案」两列），把最靠右的正文列认成题干
    if roles.get("stem") is None or roles["stem"] >= answer_index:
        stem = None
        for column in range(answer_index - 1, -1, -1):
            if _average_length(_column_values(rows_in, column)) >= 2:
                stem = column
                break
        if stem is not None and stem != roles.get("number"):
            roles["stem"] = stem
        elif roles.get("number") is not None:
            roles["stem"] = roles["number"]
            roles["number"] = None

    rest = [column for column in range(answer_index + 1, width)]
    for column in reversed(rest):
        if _small_number_like(_column_values(rows_in, column)):
            roles["score"] = column
            rest.remove(column)
            break
    for column in rest:
        if _average_length(_column_values(rows_in, column)) >= 5:
            roles["analysis"] = column
            break
    return roles


def _split_merged_options(text):
    """「A. 甲  B. 乙」→ {"A": "甲", "B": "乙"}；不是一个像样的选项串就返回 None。"""
    source = text or ""
    matches = list(RE_OPTION_MARK.finditer(source))
    if len(matches) < 2:
        return None
    letters = [match.group(1).upper() for match in matches]
    # 字母必须递增：否则是题干里恰好出现的「B.」「C.」，不是选项
    if any(letters[i] >= letters[i + 1] for i in range(len(letters) - 1)):
        return None
    options = {}
    for index, match in enumerate(matches):
        start = match.end()
        end = matches[index + 1].start() if index + 1 < len(matches) else len(source)
        body = source[start:end].strip(" 　")
        if body:
            options[letters[index]] = body
    return options or None


def _clean_type(value):
    """题型列 → 大题标题：「一、选择题（每题 2 分）」→「选择题」。"""
    text = _cell_text(value)
    text = RE_TYPE_SCORE.sub("", text)
    text = RE_TYPE_PREFIX.sub("", text)
    return text.strip(" 　、.．:：()（）")


def _guess_type(row):
    """文件没写题型时按有没有选项猜：选项题当单选，其余当解答。"""
    return "单项选择" if (row.get("options") or {}) else "简答"


CN_DIGITS = {"零": 0, "〇": 0, "一": 1, "二": 2, "两": 2, "三": 3, "四": 4,
             "五": 5, "六": 6, "七": 7, "八": 8, "九": 9}


def _cn_to_int(text):
    """「三」「十二」「二十」→ 3、12、20。老师手写的题号常常是中文。"""
    text = str(text or "").strip()
    if not text or any(char not in CN_DIGITS and char != "十" for char in text):
        return None
    if text == "十":
        return 10
    if text.startswith("十"):
        return 10 + CN_DIGITS.get(text[1:], 0)
    if "十" in text:
        head, _, tail = text.partition("十")
        return CN_DIGITS.get(head, 1) * 10 + (CN_DIGITS.get(tail, 0) if tail else 0)
    return CN_DIGITS.get(text)


def _parse_number(value):
    """题号：阿拉伯数字直接认；「三」「（十二）」这些中文题号也认（以前直接丢成 None）。"""
    text = _cell_text(value)
    match = re.search(r"\d{1,4}", text)
    if match:
        return int(match.group(0))
    return _cn_to_int(re.sub(r"[\s、.．,，)）(（:：第题-]", "", text))


def _parse_score(value):
    """(分值, 是否写了但不是数字)。老师写「2 分」也认。"""
    text = _cell_text(value)
    if not text:
        return None, False
    match = re.search(r"\d+(?:\.\d+)?", text)
    if not match:
        return None, True
    try:
        return float(match.group(0)), False
    except ValueError:
        return None, True


def _as_score(value):
    if value in (None, ""):
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def _fmt_score(value):
    if value is None:
        return ""
    if abs(value - round(value)) < 1e-6:
        return str(int(round(value)))
    return ("%.2f" % value).rstrip("0").rstrip(".")


def _cn_number(value):
    """1 → 一、10 → 十、12 → 十二。大题序号只用中文，老师看着才像卷子。

    超过 99 组时原来会 IndexError（digits[tens-1]），交给 cn_ordinal 退成阿拉伯数字。
    """
    return cn_ordinal(value)


# ---------------------------------------------------------------- 表格 → 行


def _row_from_cells(cells, roles, label, warnings):
    """一行单元格 → 一条题目；题干是空的返回 None（并记一条提示）。"""

    def cell(name):
        column = roles.get(name)
        if column is None or column >= len(cells):
            return ""
        return cells[column]

    stem = cell("stem")
    number = _parse_number(cell("number"))
    if number is None and cell("number"):
        warnings.append(_warn("number-not-readable",
                              label + "的题号没认出来（" + cell("number") + "），按顺序重排了",
                              row=label, value=cell("number")))
    if roles.get("number") is None and stem:
        # 没有题号列：题干自己带的「3、」也当题号，否则整卷题号会从 1 重排
        match = RE_STEM_NUMBER.match(stem)
        if match:
            number = int(match.group(1))
            stem = stem[match.end():]
    stem = (stem or "").strip()
    if not stem:
        warnings.append(_warn("empty-stem-row", label + "没有题干，已跳过", row=label))
        return None

    options = {}
    for letter, column in sorted((roles.get("options") or {}).items(), key=lambda item: item[1]):
        text = cells[column] if column < len(cells) else ""
        if not text:
            continue
        split = _split_merged_options(text)
        if split:
            options.update(split)       # 表头分了四列、老师却把四个选项塞进一格
        else:
            options[letter] = text
    merged_column = roles.get("merged")
    if merged_column is not None and merged_column < len(cells) and cells[merged_column]:
        split = _split_merged_options(cells[merged_column])
        if split:
            options.update(split)
        elif not options:
            options["A"] = cells[merged_column]
    options = {letter: options[letter] for letter in sorted(options)}

    answer = cell("answer").strip()
    analysis = cell("analysis").strip()
    score, bad_score = _parse_score(cell("score"))
    if not answer:
        warnings.append(_warn("answer-missing", label + "没写答案：" + stem[:20], row=label))
    if bad_score:
        warnings.append(_warn("score-not-number",
                              label + "的分值不是数字：" + cell("score"),
                              row=label, value=cell("score")))

    return {
        "number": number,
        "type": _clean_type(cell("type")),
        "material": cell("material").strip() or None,
        "stem": stem,
        "options": options,
        "answer": answer or None,
        "analysis": analysis or None,
        "score": score,
    }


def _rows_from_matrix(matrix, warnings, prefix="第 "):
    """原始单元格矩阵 → 题目行。第一行像表头就按表头映射，否则顺序兜底。"""
    rows_in = [[_cell_text(value) for value in row] for row in (matrix or [])]
    rows_in = [row for row in rows_in if any(cell for cell in row)]
    if not rows_in:
        return []
    width = max(len(row) for row in rows_in)
    roles = _header_map(rows_in[0])
    has_header = roles is not None
    if has_header:
        data = rows_in[1:]
    else:
        warnings.append(_warn("header-fallback", DEFAULT_FALLBACK_HINT))
        roles = _repair_roles(_fallback_roles(width, rows_in), rows_in, width)
        data = rows_in
    rows = []
    # 表头占掉第 1 行，报错行号也要跟着往后挪一位，否则老师按提示找不到那一行
    for offset, cells in enumerate(data, start=2 if has_header else 1):
        label = prefix + str(offset) + " 行"
        parsed = _row_from_cells(cells, roles, label, warnings)
        if parsed is not None:
            rows.append(parsed)
    return rows


# ---------------------------------------------------------------- 各格式读取


def _looks_like_bank(rows):
    """这一张表读出来的东西像不像题库：封面/说明表也能切出行，但那不是题。

    判据：过半的行至少带一个「答案 / 选项 / 分值 / 题号」——封面那种「临河三中题库」
    两行文字一条都对不上，直接跳过。
    """
    if not rows:
        return False
    good = sum(1 for row in rows
               if row.get("answer") or row.get("options") or row.get("score") or row.get("number"))
    return good * 2 >= len(rows)


def _fill_merged_down(sheet, candidate):
    """把**纵向合并**的单元格值补齐到它覆盖的每一行。

    只补跨行的合并（材料列、大题列常见），不补单行横向合并：横向合并多是
    「一、单项选择题」这种整行标题，补齐会把标题灌进每一列，反而切出怪行。
    """
    try:
        ranges = list(sheet.merged_cells.ranges)
    except Exception:
        return
    for merge in ranges:
        if merge.max_row <= merge.min_row:
            continue
        try:
            value = sheet.cell(row=merge.min_row, column=merge.min_col).value
        except Exception:
            continue
        text = _cell_text(value)
        if not text:
            continue
        for row_index in range(merge.min_row - 1, min(merge.max_row, len(candidate))):
            row = candidate[row_index]
            for column_index in range(merge.min_col, merge.max_col + 1):
                position = column_index - 1
                if position < len(row) and not _cell_text(row[position]):
                    row[position] = text


def _read_xlsx(path, warnings):
    try:
        from openpyxl import load_workbook
    except ImportError as error:  # pragma: no cover - 环境自带 openpyxl
        raise EngineError("missing-openpyxl", "缺少读 Excel 的库（openpyxl）",
                          "这个安装包不完整，重装教师助手") from error
    try:
        # 不用 read_only：只读模式拿不到 merged_cells，而老师的题库里「材料」列
        # 常被合并跨整篇短文的几行 —— 读成 None 之后，分两套时第二套那几题就没有短文了。
        book = load_workbook(path, data_only=True)
    except Exception as error:
        raise EngineError(
            "invalid-xlsx",
            "这个 Excel 打不开，可能不是 .xlsx 或者已经损坏：" + os.path.basename(path),
            "用 Excel / WPS 打开它，另存为 .xlsx 再导入",
        ) from error
    sheets = []
    try:
        # 老师的工作簿常不止一张表：封面是校名，题库在第二张，或者一个年级一张。
        # 所以每张非空表都试着读一遍，能读出题目的都收进来（不再只认第一张）。
        for sheet in book.worksheets:
            candidate = [list(row) for row in sheet.iter_rows(values_only=True)]
            if not any(any(_cell_text(cell) for cell in row) for row in candidate):
                continue
            _fill_merged_down(sheet, candidate)
            sheets.append((sheet.title or "工作表", candidate))
    finally:
        try:
            book.close()
        except Exception:
            pass
    if not sheets:
        raise EngineError("bank-empty", "这个 Excel 是空的：" + os.path.basename(path),
                          "工作表里要有表头（题干/选项A…/答案）和数据行")
    rows = []
    used = []
    for title, matrix in sheets:
        sheet_warnings = []
        got = _rows_from_matrix(matrix, sheet_warnings,
                               prefix=("「" + str(title) + "」第 ") if len(sheets) > 1 else "第 ")
        # 封面/说明那种表读出来没有像样的题干、也没有答案列，不算题库
        if not _looks_like_bank(got):
            continue
        if not any(row.get("stem") and len(row["stem"]) >= 4 for row in got):
            continue
        rows.extend(got)
        warnings.extend(sheet_warnings)
        used.append(str(title))
    if not rows:
        raise EngineError(
            "bank-empty",
            "这个 Excel 里没读出题目：" + os.path.basename(path),
            "表里要有表头（题干/选项A…/答案）和数据行；已经在里面的表：" +
            "、".join(str(title) for title, _ in sheets))
    if len(used) > 1:
        warnings.append(_warn("multi-sheet", "这个 Excel 有 " + str(len(used)) + " 张表都读出了题目，已按顺序合并："
                             + "、".join(used) + "。要只导某一张，就把它单独另存一份"))
    return rows


def _read_docx(path, warnings):
    try:
        from docx import Document
    except ImportError as error:  # pragma: no cover - 环境自带 python-docx
        raise EngineError("missing-docx", "缺少读 Word 的库（python-docx）",
                          "这个安装包不完整，重装教师助手") from error
    try:
        document = Document(path)
    except Exception as error:
        raise EngineError(
            "invalid-docx",
            "这个 Word 打不开，可能不是 .docx 或者已经损坏：" + os.path.basename(path),
            "用 Word / WPS 打开它，另存为 .docx 再导入",
        ) from error
    if not document.tables:
        raise EngineError(
            "no-table",
            "这个 Word 里没有表格：" + os.path.basename(path),
            "题库要整理成表格（一列题干、几列选项、一列答案），或者直接用 Excel",
        )
    rows = []
    for index, table in enumerate(document.tables, start=1):
        matrix = [[cell.text for cell in table_row.cells] for table_row in table.rows]
        rows.extend(_rows_from_matrix(matrix, warnings,
                                      prefix=("第 " + str(index) + " 张表第 ") if len(document.tables) > 1 else "第 "))
    return rows


def _read_text_file(path):
    """utf-8-sig 优先（Excel 导出的 csv 带 BOM），其次纯 utf-8，最后 gbk。"""
    for encoding in ("utf-8-sig", "utf-8", "gbk"):
        try:
            with open(path, "r", encoding=encoding, newline="") as handle:
                return handle.read()
        except UnicodeDecodeError:
            continue
    raise EngineError(
        "invalid-text",
        "这个文件既不是 UTF-8 也不是 GBK，中文读不出来：" + os.path.basename(path),
        "用记事本打开，另存为 UTF-8 再导入",
    )


def _guess_delimiter(text):
    """逗号还是制表符：看前几行里哪个多。整文件统计会被题干里的中文逗号带偏。"""
    lines = [line for line in (text or "").split("\n") if line.strip()][:8]
    tabs = sum(line.count("\t") for line in lines)
    commas = sum(line.count(",") for line in lines)
    return "\t" if tabs > commas else ","


def _read_csv_like(path, warnings):
    text = _read_text_file(path)
    if not text.strip():
        raise EngineError("bank-empty", "这个文件是空的：" + os.path.basename(path),
                          "第一行放表头（题干/选项A…/答案），下面放题目")
    matrix = list(csv.reader(io.StringIO(text), delimiter=_guess_delimiter(text)))
    return _rows_from_matrix(matrix, warnings)


# ---------------------------------------------------------------- 对外接口


def read_bank(path):
    """读一份题库文件 → {rows, format, warnings}。

    rows 的每一条：
      number   题号（原本没有就是 None，组卷时按顺序重排）
      type     题型（原本没有留空，组卷时按有没有选项猜）
      stem     题干（内部换行保留：老师常在题干里放一小段材料）
      options  {"A": "…", "B": "…"}，没有选项就是 {}
      answer   答案原文；没写是 None
      analysis 解析原文；没写是 None
      score    分值 float；没写是 None

    为什么「一道题都没读出来」直接报错而不是回一份空 rows：
    空着回去，后面照样能出一份只有标题的空白卷，老师以为导入成功了，
    直到打印前才发现没题 —— 不如当场拦住。
    """
    require(path not in (None, ""), "path-required", "没有给出题库文件路径",
            "把 Excel / Word 题库拖进来，或点「选择文件」")
    full = os.path.abspath(os.fspath(path))
    require(os.path.isfile(full), "file-not-found", "找不到这个题库文件：" + full,
            "确认文件没有改名、没有被移走，再选一次")

    extension = os.path.splitext(full)[1].lower()
    warnings = []
    if extension == ".xlsx":
        rows = _read_xlsx(full, warnings)
    elif extension == ".docx":
        rows = _read_docx(full, warnings)
    elif extension in (".csv", ".txt"):
        rows = _read_csv_like(full, warnings)
    elif extension == ".xls":
        raise EngineError("unsupported-format", "老式 Excel（.xls）还读不了",
                          "用 Excel / WPS 打开，另存为 .xlsx 再导入")
    else:
        raise EngineError(
            "unsupported-format",
            "这个格式的题库还读不了：" + (extension or "（没有扩展名）"),
            "目前支持 " + " / ".join(SUPPORTED),
        )

    if not rows:
        detail = "；".join(item.get("message", "") for item in warnings[:2])
        hint = (detail + "。") if detail else "题干要单独一列；没表头时按 题号/题干/A-D/答案/解析/分值 的顺序读"
        raise EngineError("bank-empty", "题库里一道题也没读出来：" + os.path.basename(full), hint)

    warnings = _dedupe_warnings(warnings)
    warnings.extend(_duplicate_warnings(rows))
    return {"rows": rows, "format": extension.lstrip("."), "warnings": warnings}


def _duplicate_warnings(rows):
    """同一道题录了两遍：不拦，但要提一句 —— 卷子上会出现两个一样的题号。"""
    index_of = {}
    duplicated = []
    for position, row in enumerate(rows, start=1):
        # 材料算进 key：两篇不同材料下各有一句「What does the underlined word mean?」
        # 是两道题，以前按题干去重会误报成重复录入。
        material = re.sub(r"\s+", "", row.get("material") or "")
        key = material + "\u0000" + re.sub(r"\s+", "", row.get("stem") or "")
        if not key:
            continue
        if key in index_of:
            duplicated.append((index_of[key], position, row["stem"][:20]))
        else:
            index_of[key] = position
    if not duplicated:
        return []
    spans = "、".join("第 %d 与第 %d 条" % (first, second) for first, second, _ in duplicated[:5])
    return [_warn("duplicate-stem",
                  "有题干重复：" + spans + "（「" + duplicated[0][2] + "」），可能重复录入了",
                  pairs=[(first, second) for first, second, _ in duplicated])]


def bank_to_text(rows, options=None):
    """题目行 → exam.parse_paper 认识的粘贴文本。

    options:
      includeAnalysis  要不要写【解析】（默认 True）
      includeScore     题干末尾要不要写（2 分）（默认 True）
      typeTitles       {"单项选择": "选择题"} 覆盖大题标题
      defaultType      题型列空着时统一当成什么

    为什么题号要跨大题连续 1..N：
    老师的 Excel 题号本来就是连续的，卷面上也按连续题号念题；
    各大题各自从 1 起会让学生在答题卡上找错行（exam 的答题卡按 1..N 编号）。
    """
    options = options or {}
    require(isinstance(rows, list), "rows-invalid",
            "题目要是一串题目（read_bank 读出来的 rows）", "先把题库导入，再生成卷子")
    include_analysis = bool(options.get("includeAnalysis", True))
    include_score = bool(options.get("includeScore", True))
    titles = options.get("typeTitles") or {}
    default_type = options.get("defaultType") or ""

    warnings = []
    groups = []
    index_of = {}
    for position, row in enumerate(rows, start=1):
        if not isinstance(row, dict):
            warnings.append(_warn("row-invalid", "第 " + str(position) + " 条不是一道题，已跳过"))
            continue
        stem = str(row.get("stem") or "").strip()
        if not stem:
            warnings.append(_warn("empty-stem", "第 " + str(position) + " 条没有题干，已跳过"))
            continue
        kind = (_clean_type(row.get("type")) or _clean_type(default_type)
                or _guess_type(row))
        if kind in index_of:
            groups[index_of[kind]][1].append((position, row, stem))
        else:
            index_of[kind] = len(groups)
            groups.append((kind, [(position, row, stem)]))

    if not groups:
        raise EngineError("rows-required", "没有可排版的题目（题库是空的）",
                          "先导入题库，或检查题干是不是写在了别的列")

    lines = []
    sections = {}
    number = 0
    last_material = None
    bare_reading = []
    for order, (kind, items) in enumerate(groups, start=1):
        lines.append(_cn_number(order) + "、" + (titles.get(kind) or kind))
        sections[kind] = len(items)
        for offset, (_position, row, stem) in enumerate(items):
            if offset:
                lines.append("")       # 空行隔开：老师看着像一份卷子，解析也不容易串行
            # 阅读/完形的短文排在这一组题目前面；同一篇连着几题只印一遍
            material = str(row.get("material") or "").strip()
            if material:
                if material != last_material:
                    lines.append("材料：")
                    lines.extend(line.rstrip() for line in material.replace("\r\n", "\n").split("\n"))
                    lines.append("")
                last_material = material
            else:
                last_material = None
                if any(word in kind.lower() for word in MATERIAL_TYPES):
                    bare_reading.append(number + 1)
            number += 1
            stem_lines = [line.rstrip() for line in stem.replace("\r\n", "\n").split("\n")]
            while stem_lines and not stem_lines[-1].strip():
                stem_lines.pop()
            head = str(number) + ". " + stem_lines[0].strip()
            body = stem_lines[1:]
            score = _as_score(row.get("score"))
            if include_score and score is not None:
                suffix = "（" + _fmt_score(score) + " 分）"
                if body:
                    body[-1] = RE_TRAILING_SCORE.sub("", body[-1]).rstrip() + suffix
                else:
                    head = RE_TRAILING_SCORE.sub("", head).rstrip() + suffix
            lines.append(head)
            lines.extend(body)

            choices = row.get("options") or {}
            if isinstance(choices, dict):
                for letter in sorted(choices):
                    text = str(choices.get(letter) or "").strip().replace("\n", " ")
                    if text:
                        lines.append(letter + ". " + text)
            else:
                warnings.append(_warn("options-invalid",
                                      "第 " + str(number) + " 题的选项格式不对，已按无选项处理"))
            if not choices and "选择" in kind:
                warnings.append(_warn("options-missing",
                                      "第 " + str(number) + " 题是选择题但没有选项"))

            answer = str(row.get("answer") or "").strip().replace("\n", " ")
            if answer:
                lines.append("【答案】" + answer)
            else:
                warnings.append(_warn("answer-missing-text",
                                      "第 " + str(number) + " 题没写答案，答案页会缺这一题"))
            if include_analysis:
                analysis = str(row.get("analysis") or "").strip().replace("\n", " ")
                if analysis:
                    # 解析必须压成一行：parse_paper 里【解析】后面的一行会被并回题干
                    lines.append("【解析】" + analysis)

    if bare_reading:
        warnings.append(_warn(
            "bare-reading",
            "第 " + "、".join(str(item) for item in bare_reading[:6])
            + (" 题" if len(bare_reading) <= 6 else " 题等 " + str(len(bare_reading)) + " 道")
            + "是阅读/完形，但题库里没有材料（短文），学生看不到文章。"
            "在题库里加一列「材料」（或「短文」），把那篇短文贴进对应的行再组一次",
        ))
    text = "\n".join(lines).rstrip() + "\n"
    return {"text": text, "count": number, "sections": sections,
            "warnings": _dedupe_warnings(warnings)}


# ---------------------------------------------------------------- 按题型挑题

RE_TYPE_NOISE = re.compile(r"[\s\u3000（）()【】\[\]、,，.。:：/\\|]+")


def _norm_type(text):
    """题型名归一：去掉空白与括号，英文小写。「单项选择」「单项选择题（2 分）」都能对上。"""
    return RE_TYPE_NOISE.sub("", str(text or "")).lower()


def _type_matches(actual, wanted):
    left, right = _norm_type(actual), _norm_type(wanted)
    if not left or not right:
        return False
    return left == right or right in left or left in right


def _type_label(row):
    name = str(row.get("type") or "").strip()
    return name or "未分类"


def bank_inspect(payload):
    """看看题库里有什么：每种题型各几道、每题几分。

    老师挑题之前得先知道题库里有多少单选、多少填空，否则「挑 10 道单选」点了才发现只有 6 道。
    payload = {path}
    """
    require(isinstance(payload, dict), "payload-invalid", "参数要是一个对象", "面板传的是 {path}")
    path = payload.get("path")
    bank = read_bank(path)
    groups, order = {}, []
    for row in bank["rows"]:
        name = _type_label(row)
        if name not in groups:
            groups[name] = {"type": name, "count": 0, "scores": []}
            order.append(name)
        groups[name]["count"] += 1
        score = row.get("score")
        if isinstance(score, (int, float)) and score > 0:
            groups[name]["scores"].append(float(score))
    types = []
    for name in order:
        group = groups[name]
        types.append({
            "type": name,
            "count": group["count"],
            "scores": sorted(set(group["scores"]))[:6],
        })
    return {
        "path": os.path.abspath(os.fspath(path)),
        "format": bank["format"],
        "count": len(bank["rows"]),
        "types": types,
        "warnings": bank["warnings"],
    }


def pick_rows(rows, pick):
    """按题型挑题。pick = {"单项选择": 10, "填空": 3}。

    返回 (挑出来的行, warnings, 每种题型实际挑到几道)。挑不满只警告不拦：
    老师更希望先拿到 8 道单选看看成品，而不是被「题库只有 8 道」顶回来。
    """
    require(isinstance(pick, (dict, list)), "pick-invalid", "挑题要写成「题型: 题数」",
            "例如 {\"单项选择\": 10, \"填空\": 3}")
    if isinstance(pick, dict):
        items = list(pick.items())
    else:
        items = []
        for item in pick:
            require(isinstance(item, dict), "pick-invalid", "挑题里每一项要是 {type, count}",
                    "例如 [{\"type\": \"单项选择\", \"count\": 10}]")
            items.append((item.get("type"), item.get("count")))
    require(bool(items), "pick-invalid", "挑题是空的", "写上「单项选择 10」这样的行，或者不填、全用")

    taken = set()
    picked, warnings, summary = [], [], {}
    available = {}
    for row in rows:
        available[_type_label(row)] = available.get(_type_label(row), 0) + 1

    for raw_type, raw_count in items:
        name = str(raw_type or "").strip()
        require(bool(name), "pick-invalid", "挑题里有一项没写题型", "一行一种题型，例如「单项选择 10」")
        try:
            count = int(float(raw_count))
        except (TypeError, ValueError):
            raise EngineError("pick-invalid", "「" + name + "」后面要跟一个题数",
                              "例如「" + name + " 10」")
        require(count > 0, "pick-invalid", "「" + name + "」的题数要大于 0", "不想要这种题型就别写这一行")

        picked_here = 0
        for index, row in enumerate(rows):
            if picked_here >= count:
                break
            if index in taken:
                continue
            if _type_matches(_type_label(row), name):
                taken.add(index)
                picked.append(row)
                picked_here += 1
        summary[name] = picked_here
        if picked_here == 0:
            hint = "、".join(sorted(available)) or "（这个题库没读出题型列）"
            raise EngineError("pick-no-type", "题库里没有「" + name + "」这种题型",
                              "题库里能挑的：" + hint)
        if picked_here < count:
            warnings.append(_warn(
                "pick-short",
                "「" + name + "」只挑到 " + str(picked_here) + " 道（你要 " + str(count) + " 道，题库里就这些）",
            ))
    return picked, _dedupe_warnings(warnings), summary


def split_sets(rows, sets):
    """把挑好的题按题型均匀分成 N 套（同一题型在每套里数量尽量一样）。

    老师的用法是「这套题库出三份，每份题量题型都差不多，别一份全是单选、一份全是阅读」，
    所以按题型分组轮流发牌，而不是按总题数从头切开。
    """
    sets = int(sets)
    if sets < 2:
        return [list(rows)]
    buckets = [list() for _ in range(sets)]
    grouped = {}
    order = []
    for index, row in enumerate(rows):
        label = _type_label(row)
        if label not in grouped:
            grouped[label] = []
            order.append(label)
        grouped[label].append((index, row))
    for offset, label in enumerate(order):
        for position, (index, row) in enumerate(grouped[label]):
            buckets[(position + offset) % sets].append((index, row))
    return [[row for _, row in sorted(bucket, key=lambda pair: pair[0])] for bucket in buckets]


def _default_title(path):
    name = os.path.splitext(os.path.basename(os.fspath(path)))[0].strip()
    return name or "测试卷"


def bank_build(payload):
    """题库 → 一整份 docx 卷子。payload = {path, title?, outputPath?, page?, options?, values?}

    options.pick = {"单项选择": 10, "填空": 3} 可以按题型挑题（不填就全用）；
    options.ab / options.teacher 和出试卷卡一致：一次出 A/B 两套，或学生卷+教师卷；
    options.sets = 3 把挑中的题按题型均匀分成 3 套，一次落 3 份卷子（返回 result["sets"]）；

    返回 exam.build_exam 的完整结果，再加一个 bank：
      bank.count     真正进卷的题数
      bank.format    源文件格式（xlsx / docx / csv / txt）
      bank.sections  {"单项选择": 5, "简答": 2}
      bank.warnings  读题库 + 排文本阶段的提示（不含组卷阶段的）

    options 同时喂给 bank_to_text 和 build_exam：两处认的键不重叠
    （includeAnalysis / scorePerQuestion 各归各的），老师不用记两套参数。
    """
    require(isinstance(payload, dict), "payload-invalid", "参数要是一个对象",
            "面板传的是 {path, title, outputPath}")
    path = payload.get("path")
    options = dict(payload.get("options") or {})
    bank = read_bank(path)
    rows = bank["rows"]
    pick_summary = None
    pick_warnings = []
    # 挑题写在顶层也认（面板与手写 HTTP 都有人这么写，静默忽略会让人以为挑题没生效）
    pick = options.pop("pick", None)
    if pick is None:
        pick = payload.get("pick")
    if pick:
        rows, pick_warnings, pick_summary = pick_rows(rows, pick)
    produced = bank_to_text(rows, options)
    title = payload.get("title") or _default_title(path)

    # 函数内 import：exam 未来若反过来引用题库（比如面板要预览），不会在模块级打死结
    from . import exam

    ab = bool(options.pop("ab", False))
    teacher = bool(options.pop("teacher", False))
    sets = options.pop("sets", None)
    if sets in (None, "", 0, "0"):
        # 与 pick 一样，顶层也认：不然老师写 {"path":…, "sets":2} 会静默只出一套
        sets = payload.get("sets")
    try:
        sets = int(sets) if sets not in (None, "", 0, "0") else 1
    except (TypeError, ValueError):
        raise EngineError("bad-sets", "「分成几套」要写一个数字",
                          "比如写 2 就是同一批题出两套（第 1 套 / 第 2 套）")
    if sets < 1:
        sets = 1
    if sets > 6:
        raise EngineError("bad-sets", "最多分成 6 套", "再多每套就没几道题了，分开建题库更合适")
    if sets > 1 and (ab or teacher):
        warnings = [{"code": "sets-ignored",
                     "message": "这次按 %d 套出；A/B 卷和教师卷要下一次单独出" % sets}]
        ab = teacher = False
    else:
        warnings = []
    if sets > len(rows):
        raise EngineError("sets-too-many", "只有 %d 道题，分不出 %d 套" % (len(rows), sets),
                          "少分几套，或者先往题库里补题")
    # A/B 两套时，默认文件名要写成「标题（A 卷）」，这样 B 卷才是「标题（B 卷）」
    # 而不是把 A 卷叫成「标题」、B 卷叫成「标题-B」（真机里两件东西看着像两份不相干的卷子）
    default_title = title + "（A 卷）" if ab else title
    output = payload.get("outputPath") or default_output("出卷", default_title, ".docx")
    exam_payload = {
        "text": produced["text"],
        # 老师（或面板）给的 overwrite 以前在这一层就丢了：目标同名时永远报 output-exists，
        # 明明说了"覆盖"也不生效。往下传给 build_exam。
        "overwrite": bool(payload.get("overwrite", False)),
        "title": title,
        "outputPath": output,
        "page": payload.get("page"),
        "values": payload.get("values") or {},
        "options": options,
    }
    if payload.get("baseTemplate"):
        # 套学校模板：题库组的卷子也要能套（以前只有「出试卷」卡认模板）
        exam_payload["baseTemplate"] = payload["baseTemplate"]
    if sets > 1:
        groups = split_sets(rows, sets)
        made = []
        first_result = None
        for index, group in enumerate(groups):
            label = "第 %d 套" % (index + 1)
            set_title = "%s %s" % (title, label)
            if payload.get("outputPath"):
                # 老师指定了输出位置：几套就落在他指定的目录里（别再甩回默认的"教师助手\出卷"）
                base_dir = os.path.dirname(os.path.abspath(output))
                stem = os.path.splitext(os.path.basename(os.path.abspath(output)))[0]
                set_output = os.path.join(base_dir, "%s %s.docx" % (stem, label))
            else:
                set_output = default_output("出卷", set_title, ".docx")
            set_text = bank_to_text(group, dict(options))
            set_payload = dict(exam_payload)
            set_payload["text"] = set_text["text"]
            set_payload["title"] = set_title
            set_payload["outputPath"] = set_output
            set_result = exam.build_exam(set_payload)
            if first_result is None:
                first_result = set_result
            made.append({
                "label": label,
                "title": set_title,
                "outputPath": set_result["outputPath"],
                "questionCount": set_result["questionCount"],
                "totalScore": set_result["totalScore"],
                "preflight": set_result["preflight"],
                "warnings": set_result.get("warnings") or [],
            })
        built = dict(first_result or {})
        built["sets"] = made
        built["multiSets"] = len(made)
        built["bank"] = {
            "count": sum(item["questionCount"] for item in made),
            "format": bank["format"],
            "sections": produced["sections"],
            "picked": pick_summary,
            "available": len(bank["rows"]),
            "warnings": bank["warnings"] + pick_warnings + produced["warnings"] + warnings,
        }
        return built
    if ab:
        if teacher:
            # A/B 与教师卷都要：四份一起出（A / B / A 教师 / B 教师）
            options["teacher"] = True
        built = exam.build_exam_pair(exam_payload)
    elif teacher:
        built = exam.build_exam_teacher(exam_payload)
    else:
        built = exam.build_exam(exam_payload)
    built["bank"] = {
        "count": produced["count"],
        "format": bank["format"],
        "sections": produced["sections"],
        "picked": pick_summary,
        "available": len(bank["rows"]),
        "warnings": bank["warnings"] + pick_warnings + produced["warnings"] + warnings,
    }
    return built

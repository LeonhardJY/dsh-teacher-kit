"""占位符：把"这一份卷子"变成"下学期还能用的模板"，以及造新卷时把值填回去。

两类占位符要区别对待：
- {{schoolYear}} {{examTitle}} {{totalScore}} {{minutes}} {{className}} {{teacher}} {{blank}} —— 纯文本，替换即可
- {{page}} {{pages}} —— 必须是真页码域（PAGE / NUMPAGES），否则导出的 PDF 里是两片死文字
"""

import copy
import re

from docx.oxml import OxmlElement
from docx.oxml.ns import qn

TOKEN_RE = re.compile(r"\{\{\s*([A-Za-z][A-Za-z0-9_]*)\s*\}\}")
# pages = 全文总页数（NUMPAGES）；sectionPages = 本节页数（SECTIONPAGES）。
# 多纸型混排（正文 8 开 + 答题卡 A4）时，答题卡那几页要从 1 重新编、总数只算本节，
# 否则「共 8 页」会把 A4 答题卡也数进 8 开卷子里。
FIELD_TOKENS = {"page": " PAGE ", "pages": " NUMPAGES ", "sectionpages": " SECTIONPAGES "}

# 老师（和学校模板）写页脚时用的是各种叫法：「共 {{total}} 页」「共{{numPages}}页」。
# 只认 {{pages}} 的话，这些写法会原样留在页脚里（印出来是「共{{total}}页」），
# 印前检查还会判 fail、连 PDF 都导不出。这里统一收口成上面三种真域。
TOKEN_ALIASES = {
    "total": "pages",
    "totalpages": "pages",
    "numpages": "pages",
    "pagecount": "pages",
    "总页数": "pages",
    "sectionpages": "sectionpages",
    "sectionpagecount": "sectionpages",
    "page": "page",
    "pageno": "page",
    "pagenumber": "page",
    "页码": "page",
}


def canonical_token(name):
    """把 {{total}} 这类别名收成 {{pages}}；认不出就原样返回。"""
    key = (name or "").strip().lower()
    return TOKEN_ALIASES.get(key, name)


def tokens_in(text):
    return [canonical_token(match.group(1)) for match in TOKEN_RE.finditer(text or "")]


def _base_rpr(paragraph):
    for run in paragraph.runs:
        rpr = run._element.find(qn("w:rPr"))
        if rpr is not None:
            return copy.deepcopy(rpr)
    return None


def _wipe(paragraph):
    """清掉段落里的所有文字与域，只留段落属性（对齐、缩进、样式）。"""
    for child in list(paragraph._p):
        if child.tag in (qn("w:r"), qn("w:fldSimple"), qn("w:hyperlink")):
            paragraph._p.remove(child)


def _new_run(paragraph, text, rpr=None):
    run = paragraph.add_run(text)
    if rpr is not None:
        element = run._element
        existing = element.find(qn("w:rPr"))
        if existing is not None:
            element.remove(existing)
        element.insert(0, copy.deepcopy(rpr))
    return run


def _new_field(paragraph, instr, sample="1", rpr=None):
    """插一个真域：LibreOffice/Word 打开时自动计算，PDF 出来才是"第3页 共6页"。"""
    field = OxmlElement("w:fldSimple")
    field.set(qn("w:instr"), instr)
    run = OxmlElement("w:r")
    if rpr is not None:
        run.append(copy.deepcopy(rpr))
    text = OxmlElement("w:t")
    text.text = sample
    run.append(text)
    field.append(run)
    paragraph._p.append(field)
    return field


def _is_field_run(element):
    """这个 run 是域的一部分（fldChar / instrText）：动它会把页码域拆坏。"""
    return element.find(qn("w:fldChar")) is not None or element.find(qn("w:instrText")) is not None


def _run_text(element):
    return "".join(node.text or "" for node in element.iter(qn("w:t")))


def _run_with_text(text, rpr=None):
    run = OxmlElement("w:r")
    if rpr is not None:
        run.append(copy.deepcopy(rpr))
    node = OxmlElement("w:t")
    node.set(qn("xml:space"), "preserve")
    node.text = text
    run.append(node)
    return run


def _field_with_text(instr, rpr=None, sample="1"):
    field = OxmlElement("w:fldSimple")
    field.set(qn("w:instr"), instr)
    run = OxmlElement("w:r")
    if rpr is not None:
        run.append(copy.deepcopy(rpr))
    node = OxmlElement("w:t")
    node.text = sample
    run.append(node)
    field.append(run)
    return field


def _pieces_for(text, values, want_fields, unfilled):
    """一段文字 → [("text", 文字) | ("field", 域指令)]，没值可填的占位符原样留着。"""
    pieces = []
    cursor = 0
    for match in TOKEN_RE.finditer(text):
        if match.start() > cursor:
            pieces.append(("text", text[cursor:match.start()]))
        cursor = match.end()
        name = match.group(1)
        key = canonical_token(name).lower()
        filled_value = values.get(name, values.get(key))
        if key in FIELD_TOKENS and want_fields:
            pieces.append(("field", FIELD_TOKENS[key]))
        elif filled_value is not None and str(filled_value) != "":
            pieces.append(("text", _expand(str(filled_value), values)))
        else:
            unfilled.append(name)
            pieces.append(("text", match.group(0)))
    if cursor < len(text):
        pieces.append(("text", text[cursor:]))
    return pieces


def _expand(value, values, depth=0):
    """值里还带占位符时再展开一层（「使用班级：{{className}}」这种嵌套是常见写法）。

    以前只替换一趟：值里剩下的 {{blank}} 就原样印到纸上，
    页脚出现「使用班级：使用班级：{{blank}}」。最多展开 4 层，防自引用死循环。
    """
    if depth >= 4 or "{{" not in value:
        return value
    pieces = _pieces_for(value, values, False, [])
    out = "".join(payload for kind, payload in pieces if kind == "text")
    if out == value:
        return value
    return _expand(out, values, depth + 1)


def fill_paragraph(paragraph, values, want_fields=True):
    """把段落里的占位符换成值，返回没被填上的占位符名。

    只重建「纯文字 run」：页码域（fldSimple，或 fldChar + instrText）原样留着 ——
    老写法整段清空重建，会把「第 {{page}} 页」里那个真的页码域一起删掉，
    导出的卷子页脚就变成「第  页 共  页」。
    """
    element = paragraph._p
    plain = [child for child in element if child.tag == qn("w:r") and not _is_field_run(child)]
    if "{{" not in "".join(_run_text(child) for child in plain):
        return []

    groups, current = [], []
    for child in list(element):
        if child.tag == qn("w:r") and not _is_field_run(child):
            current.append(child)
            continue
        if current:
            groups.append(current)
            current = []
    if current:
        groups.append(current)

    unfilled = []
    for group in groups:
        text = "".join(_run_text(run) for run in group)
        if "{{" not in text:
            continue
        pieces = _pieces_for(text, values, want_fields, unfilled)
        anchor = group[0]
        rpr = anchor.find(qn("w:rPr"))
        rpr = copy.deepcopy(rpr) if rpr is not None else None
        index = list(element).index(anchor)
        for run in group:
            element.remove(run)
        insert_at = index
        for kind, payload in pieces:
            if kind == "text":
                if not payload:
                    continue
                node = _run_with_text(payload, rpr)
            else:
                node = _field_with_text(payload, rpr)
            element.insert(insert_at, node)
            insert_at += 1
    return unfilled


def _story_parts(document):
    parts = []
    for section in document.sections:
        for name in ("header", "first_page_header", "even_page_header",
                     "footer", "first_page_footer", "even_page_footer"):
            part = getattr(section, name, None)
            if part is not None:
                parts.append(part)
    return parts


def _paragraphs_of(container):
    found = list(getattr(container, "paragraphs", []) or [])
    for table in getattr(container, "tables", []) or []:
        for row in table.rows:
            for cell in row.cells:
                found.extend(_paragraphs_of(cell))
    return found


def fill_document(document, values, include_body=True):
    """页眉/页脚/正文里的占位符全部填掉，返回填了多少段、还剩哪些没填。"""
    values = {key: value for key, value in (values or {}).items() if value is not None}
    touched = 0
    unfilled = set()
    containers = _story_parts(document)
    if include_body:
        containers.append(document)
    for container in containers:
        for paragraph in _paragraphs_of(container):
            if "{{" not in (paragraph.text or ""):
                continue
            missing = fill_paragraph(paragraph, values)
            touched += 1
            if missing:
                unfilled.update(missing)
    return {"filled": touched, "unfilled": sorted(unfilled)}


def replace_text(paragraph, text):
    """把整段换成一段文字（保留第一个 run 的字体格式）。"""
    rpr = _base_rpr(paragraph)
    _wipe(paragraph)
    return _new_run(paragraph, text, rpr)

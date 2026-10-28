"""教育局/学校公文流转：收文登记 → 拟办意见 → 领导批示 → 承办流转 → 回复稿。

边界
----
* 只是**辅助**：这几张单子排得清楚、要素齐全、能直接印出来送签，但签不签、怎么批、
  谁承办，都是人和制度的事，工具不替任何一位领导拿主意，也不生成"已办结"这种结论。
* 不联网、不发公文、不伪造红头与公章。红头只在老师明确要、并给了单位全称时才排，
  印章位置留空（真实公文必须盖章，工具画一枚假章是给学校添麻烦）。
* 版式按 GB/T 9704 的常规做法：正文三号仿宋、行距 28pt、首行缩进 2 字符、
  标题二号黑体加粗（小标宋不一定装，黑体一定在，印出来同样是公文相）、一级标题「一、」
  黑体三号、主送机关顶格、发文机关署名右空二字、成文日期右空四字、落款右对齐且日期另起一行、
  A4 上 37 下 35 左 28 右 26mm。表格按内容分配列宽（登记表不能四等分）、
  单元格左右留 0.28cm、表头黑体加粗浅底。
  学校有自己的公文模板时（payload.baseTemplate）**优先用模板的版式**，这里只填内容。

产出
----
一个 .docx，按办理顺序装订：①收文登记表 ②来文要点与要求 ③拟办意见 ④领导批示栏
⑤承办流转单 ⑥回复稿（给要求书面回复的来文）。老师可以只印其中几页。
"""

import datetime
import os

from .docx_build import build_docx, set_run_font
from .errors import require
from .profile import cn_ordinal
from .paths import default_output

# GB/T 9704 的常用页面：A4，上 37 下 35 左 28 右 26（单位毫米）
GB_MARGINS = {"top": 37, "bottom": 35, "left": 28, "right": 26}

BODY_FONT = "仿宋"          # 公文正文
TITLE_FONT = "黑体"         # 二号标题（小标宋不一定装，黑体几乎每台机器都有）
HEAD_FONT = "宋体"          # 红头「XX文件」——宋体加粗更接近小标宋的样子
BODY_SIZE = 16.0            # 三号
LINE_SPACING = 28.0         # 公文常见 28pt 固定行距
TABLE_FONT = "仿宋"         # 表内文字（比正文小一号，格子才装得下）
TABLE_SIZE = 12.0           # 小四
TABLE_HEAD_FONT = "黑体"
CELL_MARGIN_TWIPS = 158     # 单元格左右内边距 ≈ 0.28cm
CELL_PAD_TWIPS = 57         # 上下内边距 ≈ 0.1cm

URGENCY_LEVELS = ("普通", "急件", "加急", "特急")

REQUIRED_FIELDS = (
    ("from", "来文单位"),
    ("title", "文件标题"),
    ("docNo", "发文字号"),
    ("receivedAt", "收文日期"),
    ("dept", "承办部门"),
)


def _block_text(value):
    """把回复稿里的一段变成纯文本。

    以前直接 str(value)：老师写了 {"type":"text","text":"…"} 时，成品正文里就原样
    印出一行 Python 字典字面量 —— 一份公文中间冒出 {'type': 'text'…}。
    """
    if isinstance(value, str):
        return value.strip()
    if isinstance(value, dict):
        for key in ("text", "heading", "lead", "content"):
            text = value.get(key)
            if isinstance(text, str) and text.strip():
                return text.strip()
        raise EngineError(
            "block-invalid",
            "回复稿里有一段是对象，但没写正文：" + "、".join(str(key) for key in value)[:60],
            "写成 {\"type\": \"text\", \"text\": \"第一段正文\"}，或者直接写一段字符串")
    if value in (None, ""):
        return ""
    raise EngineError(
        "block-invalid",
        "回复稿里的段落只能是字符串或 {type, text} 对象，这次收到的是 " + type(value).__name__,
        "把这一段改成纯文本，例如「现将有关情况报告如下：」")


def _text(value):
    return str(value).strip() if value not in (None, "") else ""


def _paragraph(text, size=BODY_SIZE, font=BODY_FONT, align=None, bold=False,
               indent=True, spacing=LINE_SPACING, after=0, color=None):
    """一个公文正文段落：三号仿宋、行距 28、首行缩进两字（标题/主送机关不缩进）。"""
    run = {"text": text, "sizePt": size, "font": font}
    if bold:
        run["bold"] = True
    if color:
        run["color"] = color
    # sizePt 也放在段落上：docx_build 按「首行缩进 2 字符 × 本段字号」换算缩进，
    # 只写在 run 上它会一律按 12pt 算成 24pt，三号正文该是 32pt。
    block = {"type": "paragraph", "runs": [run], "sizePt": size,
             "lineSpacingPt": spacing, "spaceAfterPt": after}
    if align:
        block["align"] = align
    if indent:
        block["firstLineIndentChars"] = 2
    return block


def _section(text, after=2):
    """一级标题「一、」：黑体三号顶格。用黑体而不是把仿宋加粗——加粗的仿宋在纸上几乎看不出层次。"""
    return _paragraph(text, size=BODY_SIZE, font=TITLE_FONT, bold=True, indent=False,
                      spacing=LINE_SPACING, after=after)


def _label_row(label, value):
    """登记表的一行：左格项目名、右格内容。"""
    return [label, value if value not in (None, "") else "　"]


def _registry_rows(doc):
    fields = doc["fields"]
    rows = [["收文登记表", ""]]
    rows.append(_label_row("来文单位", fields["from"]))
    rows.append(_label_row("发文字号", fields["docNo"]))
    rows.append(_label_row("文件标题", fields["title"]))
    rows.append(_label_row("收文日期", fields["receivedAt"]))
    rows.append(_label_row("份数", fields["copies"]))
    rows.append(_label_row("密级", fields["secret"]))
    rows.append(_label_row("紧急程度", fields["urgency"]))
    rows.append(_label_row("承办部门", fields["dept"]))
    rows.append(_label_row("办理期限", fields["deadline"]))
    rows.append(_label_row("收文登记人", fields["handler"]))
    rows.append(_label_row("收文编号", fields["registryNo"]))
    return rows


def _steps_rows(doc):
    steps = doc["steps"]
    rows = [["序", "办理环节", "承办部门", "责任人", "时限", "办理情况"]]
    for index, step in enumerate(steps):
        rows.append([
            str(index + 1),
            _text(step.get("name")) or _text(step.get("title")) or _text(step.get("stage")) or "　",
            _text(step.get("dept")) or _text(step.get("department")) or "　",
            _text(step.get("owner")) or _text(step.get("person")) or "　",
            _text(step.get("due")) or _text(step.get("deadline")) or _text(step.get("limit")) or "　",
            _text(step.get("status")) or _text(step.get("note")) or _text(step.get("body"))
            or _text(step.get("result")) or "　",
        ])
    return rows


TABLE_HEAD_FILL = "F2F4F7"    # 表头浅底：打印出来是浅灰，不抢字
TABLE_LABEL_FILL = "F6F7F9"   # 登记表项目名列的浅底


def _char_width(text):
    """按"看起来占几格"算宽度：汉字 1、西文数字 0.55。列宽照这个分，不能按 len 分。"""
    total = 0.0
    for char in str(text or ""):
        total += 1.0 if ord(char) > 127 else 0.55
    return total


def _polish_tables(path):
    """表格按内容分列宽 + 压内边距 + 表头浅底（docx_build 只会等分，改不得——试卷听写单都在用）。

    公文登记表要是四等分，「文件标题」那格只剩几个字，一长串标题挤成好几行；
    流转单里「序」「时限」也不该跟「办理环节」一样宽。所以产出后重开文件单独调：
    gridCol 与单元格宽度按内容权重分、单元格左右 0.28cm、表头黑体加粗浅底、长行不跨页。
    调不好也不抛错：表格没排漂亮不该让整份单子出不来。
    """
    try:
        from docx import Document
        from docx.enum.table import WD_ALIGN_VERTICAL
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        from docx.oxml.ns import qn
        from docx.shared import Twips
    except Exception:  # 没装 python-docx 就保持 build_docx 的原样
        return False
    try:
        document = Document(path)
    except Exception:
        return False

    def shade(cell, fill):
        tc_pr = cell._tc.get_or_add_tcPr()
        node = tc_pr.find(qn("w:shd"))
        if node is None:
            node = tc_pr.makeelement(qn("w:shd"), {})
            tc_pr.append(node)
        node.set(qn("w:val"), "clear")
        node.set(qn("w:color"), "auto")
        node.set(qn("w:fill"), fill)

    def unique_cells(row):
        cells, seen = [], set()
        for cell in row.cells:
            if id(cell._tc) in seen:
                continue
            seen.add(id(cell._tc))
            cells.append(cell)
        return cells

    def row_flag(row, tag):
        tr_pr = row._tr.get_or_add_trPr()
        if tr_pr.find(qn(tag)) is None:
            tr_pr.append(tr_pr.makeelement(qn(tag), {}))

    def set_margins(table):
        tbl_pr = table._tbl.tblPr
        node = tbl_pr.find(qn("w:tblCellMar"))
        if node is None:
            node = tbl_pr.makeelement(qn("w:tblCellMar"), {})
            look = tbl_pr.find(qn("w:tblLook"))
            if look is not None:
                look.addprevious(node)   # 按 schema 顺序插在 tblLook 之前
            else:
                tbl_pr.append(node)
        for tag, value in (("w:top", CELL_PAD_TWIPS), ("w:left", CELL_MARGIN_TWIPS),
                           ("w:bottom", CELL_PAD_TWIPS), ("w:right", CELL_MARGIN_TWIPS)):
            child = node.find(qn(tag))
            if child is None:
                child = node.makeelement(qn(tag), {})
                node.append(child)
            child.set(qn("w:w"), str(value))
            child.set(qn("w:type"), "dxa")

    touched = False
    for table in document.tables:
        rows = list(table.rows)
        if len(rows) < 2:
            continue
        grid = table._tbl.tblGrid
        columns = grid.findall(qn("w:gridCol")) if grid is not None else []
        if len(columns) < 2:
            continue
        total = 0
        for column in columns:
            try:
                total += int(column.get(qn("w:w")) or 0)
            except (TypeError, ValueError):
                pass
        if total <= 0:
            continue

        # 首行左格有字、右格空 → 是跨列表题（「收文登记表」），先合并再分列宽
        title_row = False
        head_cells = unique_cells(rows[0])
        if len(columns) == 2 and len(head_cells) == 2 and head_cells[0].text.strip()                 and not head_cells[1].text.strip():
            merged = head_cells[0].merge(head_cells[1])
            title_row = True
            # 合并会把两个格子的段落都留下，第二段空行会把标题行撑高一整行
            for extra in merged.paragraphs[1:]:
                if not extra.text.strip():
                    extra._element.getparent().remove(extra._element)
            for paragraph in merged.paragraphs:
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in paragraph.runs:
                    set_run_font(run, TABLE_HEAD_FONT, 14.0)
                    run.bold = True
            shade(merged, TABLE_HEAD_FILL)

        # 列宽权重 = 这一列最长的那个格子；跨列的格子不参与（一句话会把两列撑平）
        weights = []
        for index in range(len(columns)):
            longest = 3.0
            for row in rows:
                cells = row.cells
                if index >= len(cells):
                    continue
                cell = cells[index]
                if any(cells[other]._tc is cell._tc for other in range(index)):
                    continue
                longest = max(longest, _char_width(cell.text))
            weights.append(min(longest, 18.0))
        total_weight = sum(weights) or 1.0
        widths = [max(int(total * weight / total_weight), int(total * 0.04)) for weight in weights]
        widths[-1] += total - sum(widths)   # 取整误差落到最后一列
        if widths[-1] <= 0:
            widths = [total // len(columns)] * len(columns)
            widths[-1] += total - sum(widths)
        for column, width in zip(columns, widths):
            column.set(qn("w:w"), str(width))
        tbl_w = table._tbl.tblPr.find(qn("w:tblW"))
        if tbl_w is not None:
            tbl_w.set(qn("w:type"), "dxa")
            tbl_w.set(qn("w:w"), str(total))

        set_margins(table)
        touched = True
        for r_index, row in enumerate(rows):
            row_flag(row, "w:cantSplit")   # 一行别被页断成两半
            seen = set()
            for index, cell in enumerate(row.cells):
                if id(cell._tc) in seen:
                    continue
                seen.add(id(cell._tc))
                span = 1
                tc_pr = cell._tc.tcPr
                if tc_pr is not None:
                    node = tc_pr.find(qn("w:gridSpan"))
                    if node is not None:
                        try:
                            span = max(1, int(node.get(qn("w:val")) or 1))
                        except (TypeError, ValueError):
                            span = 1
                try:
                    cell.width = Twips(sum(widths[index:index + span]) or total)
                except Exception:
                    pass
                try:
                    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                except Exception:
                    pass
                if title_row and r_index == 0:
                    continue          # 跨列表题那格已经单独排过
                is_header = (not title_row) and r_index == 0
                is_label = title_row and index == 0
                for paragraph in cell.paragraphs:
                    if is_header or is_label or _char_width(cell.text) <= 6:
                        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    for run in paragraph.runs:
                        if is_header or is_label:
                            set_run_font(run, TABLE_HEAD_FONT, TABLE_SIZE)
                            run.bold = True
                        else:
                            set_run_font(run, TABLE_FONT, TABLE_SIZE)
                if is_header:
                    shade(cell, TABLE_HEAD_FILL)
                elif is_label:
                    shade(cell, TABLE_LABEL_FILL)
        if not title_row:
            row_flag(rows[0], "w:tblHeader")   # 表头跨页重复

    if not touched:
        return False
    try:
        document.save(path)
    except Exception:
        return False
    return True


def _check(doc):
    """把"送签之前会被打回来"的缺项列出来。不拦生成——老师可能就要先印个空表。"""
    warnings = []
    missing = []
    fields = doc["fields"]
    for key, label in REQUIRED_FIELDS:
        if not _text(fields.get(key)):
            missing.append(label)
    if missing:
        level = "error" if len(missing) >= 3 else "warn"
        warnings.append({"code": "field-missing", "level": level,
                         "message": "这几项还空着：" + "、".join(missing) + "（登记表上空着送签，办公室会退回来）",
                         "fields": missing})
    if _text(doc.get("urgencyInput")) and _text(doc.get("urgencyInput")) != _text(fields.get("urgency")):
        warnings.append({"code": "urgency-normalized", "level": "info",
                         "message": "「" + _text(doc.get("urgencyInput")) + "」不是标准紧急程度（平件/普通/急件/加急/特急），"
                                    "登记表按「" + _text(fields.get("urgency")) + "」印了"})
    if _text(fields.get("urgency")) in ("急件", "加急", "特急") and not _text(fields.get("deadline")):
        warnings.append({"code": "urgent-no-deadline", "level": "error",
                         "message": "标了" + _text(fields.get("urgency")) + "却没写办理期限——急件必须先落到具体日期"})
    opinion = _text(doc.get("opinion"))
    if not opinion:
        warnings.append({"code": "opinion-empty", "level": "warn",
                         "message": "还没有拟办意见：签发前要有人先写「拟请教务处牵头，XX 日前报回」"})
    elif len(opinion) < 12:
        warnings.append({"code": "opinion-short", "level": "info",
                         "message": "拟办意见只有 " + str(len(opinion)) + " 字，通常要写清谁牵头、怎么办、什么时候报回"})
    if doc["steps"]:
        no_due = [step for step in doc["steps"] if not _text(step.get("due"))]
        if no_due:
            warnings.append({"code": "step-no-due", "level": "info",
                             "message": "流转单里有 " + str(len(no_due)) + " 个环节没写时限，办公室催办时没法跟"})
    else:
        warnings.append({"code": "steps-empty", "level": "info",
                         "message": "还没排承办流转环节（登记 → 拟办 → 批示 → 承办 → 归档，通常 4—5 步）"})
    reply = doc.get("reply") or {}
    if reply:
        if not _text(reply.get("to")):
            warnings.append({"code": "reply-no-to", "level": "warn",
                             "message": "回复稿缺主送机关（就是顶格写「XX 教育局：」的那一行）"})
        if not _text(reply.get("date")) and not _text(reply.get("signature")):
            warnings.append({"code": "reply-no-sign", "level": "warn",
                             "message": "回复稿没有落款单位和成文日期，打印出来没法盖章报送"})
        body_text = " ".join(_text(block) for block in (reply.get("blocks") or []))
        if len(body_text) < 40:
            warnings.append({"code": "reply-too-short", "level": "info",
                             "message": "回复稿正文只有 " + str(len(body_text)) + " 字，通常要说清收到什么、怎么办的、什么结果"})
    return warnings, {"missing": missing}


def docflow_build(payload):
    """出一份公文办理单（可含回复稿）。"""
    title = _text(payload.get("title"))
    require(bool(title), "title-required", "还没有文件标题：来文标题照抄原文，别自己概括（办公室按标题归档）", "把来文的标题整句抄进来，比如「关于开展校园安全检查的通知」")
    unit = _text(payload.get("unit"))
    options = payload.get("options") or {}
    fields = {
        "from": _text(payload.get("from")),
        "docNo": _text(payload.get("docNo")),
        "title": title,
        "receivedAt": _text(payload.get("receivedAt")) or datetime.date.today().strftime("%Y-%m-%d"),
        "copies": _text(payload.get("copies")) or "1",
        "secret": _text(payload.get("secret")) or "内部",
        "urgency": _text(payload.get("urgency")) or "普通",
        "dept": _text(payload.get("dept")),
        "deadline": _text(payload.get("deadline")),
        "handler": _text(payload.get("handler")),
        "registryNo": _text(payload.get("registryNo")),
    }
    urgency = fields["urgency"]
    urgency_input = _text(payload.get("urgency"))
    if urgency not in URGENCY_LEVELS:
        urgency = "普通"
        fields["urgency"] = urgency
    steps = payload.get("steps") or []
    if isinstance(steps, (str, dict)):
        steps = []
    steps = [step for step in steps if isinstance(step, dict)]
    doc = {
        "fields": fields,
        "urgencyInput": urgency_input,
        "unit": unit,
        "summary": _text(payload.get("summary")),
        "items": [_text(item) for item in (payload.get("items") or []) if _text(item)],
        "opinion": _text(payload.get("opinion")),
        "leaderNote": _text(payload.get("leaderNote")),
        "steps": steps,
        "reply": payload.get("reply") if isinstance(payload.get("reply"), dict) else {},
    }
    warnings, extra = _check(doc)

    blocks = []
    # 红头：只在老师明确要、且知道单位全称时排；印章位留空
    if options.get("redHead") and unit:
        blocks.append({"type": "paragraph", "align": "center", "spaceAfterPt": 2,
                       "runs": [{"text": unit + "文件", "sizePt": 22, "font": HEAD_FONT,
                                 "bold": True, "color": "C00000"}]})
    if unit and not options.get("redHead"):
        # 学校画像里记了本单位：要在纸面上看得见（以前只进页脚，老师以为画像没生效）
        blocks.append({"type": "paragraph", "align": "center", "spaceAfterPt": 2,
                       "runs": [{"text": unit, "sizePt": 12, "font": BODY_FONT}]})
    blocks.append({"type": "paragraph", "align": "center", "spaceAfterPt": 6,
                   "runs": [{"text": "公文办理单", "sizePt": 22, "font": TITLE_FONT, "bold": True}]})
    blocks.append({"type": "paragraph", "align": "center", "spaceAfterPt": 10,
                   "runs": [{"text": "（收文登记 · 拟办 · 批示 · 流转）", "sizePt": 12, "font": BODY_FONT}]})
    blocks.append({"type": "table", "tableStyle": "Table Grid", "rows": _registry_rows(doc)})

    # 四节先按顺序拼成列表再编号：以前「一、来文要点」没内容就不印，后面却写死「二/三/四」，
    # 公文直接从「二、拟办意见」开始 —— 跳号了。
    head_blocks = []
    if doc["summary"]:
        head_blocks.append(_paragraph(doc["summary"]))
    for index, item in enumerate(doc["items"]):
        head_blocks.append(_paragraph("（" + cn_ordinal(index + 1) + "）" + item))
    opinion_blocks = [_paragraph(doc["opinion"] or "（请填写：谁牵头办理、怎么办、什么时候报回）", after=8)]
    leader_blocks = [
        _paragraph(doc["leaderNote"] or "（待批示）", after=8),
        _paragraph("批示人：　　　　　　批示日期：　　年　　月　　日", size=12, indent=False,
                   spacing=22, after=6),
    ]
    steps_blocks = []
    if doc["steps"]:
        steps_blocks.append({"type": "table", "tableStyle": "Table Grid", "rows": _steps_rows(doc)})
    else:
        steps_blocks.append({"type": "table", "tableStyle": "Table Grid",
                             "rows": [["序", "办理环节", "承办部门", "责任人", "时限", "办理情况"]]})
    steps_blocks.append(_paragraph("备注：办结后将本单与办理结果一并归档。", size=12, indent=False,
                                   spacing=22, after=0))
    sections_build = []
    if head_blocks:
        sections_build.append(("来文要点", head_blocks, False))
    # 拟办意见单独起页：签发时这一页要签字，跟登记表挤在一张纸上不方便
    sections_build.append(("拟办意见", opinion_blocks, True))
    sections_build.append(("领导批示", leader_blocks, False))
    sections_build.append(("承办流转单", steps_blocks, False))
    for index, section_item in enumerate(sections_build):
        title_text, inner, break_before = section_item
        if break_before:
            blocks.append({"type": "pageBreak"})
        blocks.append(_section(cn_ordinal(index + 1) + "、" + title_text))
        blocks.extend(inner)

    reply = doc["reply"]
    if reply:
        blocks.append({"type": "pageBreak"})
        reply_title = _text(reply.get("title")) or ("关于落实《" + title + "》情况的报告")
        blocks.append({"type": "paragraph", "align": "center", "spaceAfterPt": 12,
                       "runs": [{"text": reply_title, "sizePt": 22, "font": TITLE_FONT, "bold": True}]})
        if _text(reply.get("to")):
            blocks.append(_paragraph(_text(reply["to"]).rstrip("：:") + "：", indent=False, after=0))
        for text in reply.get("blocks") or []:
            text = _block_text(text)
            if text:
                blocks.append(_paragraph(text))
        sign = _text(reply.get("signature")) or unit
        date = _text(reply.get("date")) or datetime.date.today().strftime("%Y年%m月%d日").replace("年0", "年").replace("月0", "月")
        blocks.append(_paragraph("　", spacing=22, after=0))
        if sign:
            blocks.append({"type": "paragraph", "align": "right", "spaceAfterPt": 0,
                           "lineSpacingPt": LINE_SPACING,
                           "runs": [{"text": sign + "　", "sizePt": BODY_SIZE, "font": BODY_FONT}]})
        blocks.append({"type": "paragraph", "align": "right", "spaceAfterPt": 0,
                       "lineSpacingPt": LINE_SPACING,
                       "runs": [{"text": date + "　　　　", "sizePt": BODY_SIZE, "font": BODY_FONT}]})
        if options.get("sealNote", True):
            blocks.append(_paragraph("（盖章处）", size=10.5, align="right", indent=False, spacing=18, after=0,
                                     color="808080"))

    output = payload.get("outputPath")
    if not output:
        stem = title if len(title) <= 24 else title[:24]
        output = default_output("公文", stem + "-办理单", ".docx")
    spec = {
        "outputPath": output,
        "page": {"size": "A4", "orientation": "portrait", "marginsMm": GB_MARGINS, "columns": 1},
        "blocks": blocks,
        "properties": {"title": title + " 公文办理单", "author": unit or "教师助手"},
        "footer": "第 {{page}} 页 共 {{pages}} 页　　" + (unit or "") + "公文办理单",
        "overwrite": bool(payload.get("overwrite", False)),
    }
    if _text(payload.get("baseTemplate")):
        template = os.path.abspath(_text(payload.get("baseTemplate")))
        require(os.path.isfile(template), "template-missing", "找不到这份公文模板：" + template)
        spec["baseTemplate"] = template
        spec.pop("page", None)  # 用学校模板的版式，不拿国标页边距硬盖
    built = build_docx(spec)
    _polish_tables(built["outputPath"])
    return {
        "outputPath": built["outputPath"],
        "bytes": built["bytes"],
        "title": title,
        "unit": unit,
        "fields": fields,
        "sections": ["收文登记表", "来文要点", "拟办意见", "领导批示", "承办流转单"] + (["回复稿"] if reply else []),
        "steps": len(steps),
        "hasReply": bool(reply),
        "warnings": warnings + list(built.get("warnings") or []),
        **extra,
    }

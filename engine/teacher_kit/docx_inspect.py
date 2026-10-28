"""读 docx 的"版式事实"：页面、页边距、页眉页脚、字体、样式、结构。

这些数字是后面两件事的输入：
- 模板反推（学校模板到底怎么排的）
- 印前检查（生成的卷子有没有偏离模板）
"""

from docx import Document
from docx.shared import Emu

from .errors import EngineError, require
from .units import emu_to_mm, pt_to_mm


def _section_facts(section):
    """一节（通常是"一页规则"）的页面事实。"""
    header_text = "\n".join(p.text for p in section.header.paragraphs if p.text.strip())
    footer_text = "\n".join(p.text for p in section.footer.paragraphs if p.text.strip())
    return {
        "orientation": "横向" if section.orientation == 1 else "纵向",
        "pageWidthMm": emu_to_mm(section.page_width),
        "pageHeightMm": emu_to_mm(section.page_height),
        "marginTopMm": emu_to_mm(section.top_margin),
        "marginBottomMm": emu_to_mm(section.bottom_margin),
        "marginLeftMm": emu_to_mm(section.left_margin),
        "marginRightMm": emu_to_mm(section.right_margin),
        "headerDistanceMm": emu_to_mm(section.header_distance),
        "footerDistanceMm": emu_to_mm(section.footer_distance),
        "headerText": header_text or None,
        "footerText": footer_text or None,
        "columns": _column_count(section),
    }


def _column_count(section):
    """分栏数（8 开试卷常分两栏）。python-docx 没暴露，读底层 XML。"""
    try:
        cols = section._sectPr.find(
            "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}cols"
        )
        if cols is None:
            return 1
        return int(cols.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}num") or 1)
    except Exception:
        return 1


def _run_fonts(document):
    """文档里实际出现过的字体（按出现次数排序）。中文字体要单独看 eastAsia。"""
    counts = {}
    paragraphs = list(document.paragraphs)
    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                paragraphs.extend(cell.paragraphs)
    for paragraph in paragraphs:
        for run in paragraph.runs:
            name = run.font.name
            if not name:
                try:
                    rpr = run._element.rPr
                    east = rpr.rFonts.get(
                        "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}eastAsia"
                    ) if rpr is not None and rpr.rFonts is not None else None
                    name = east
                    if not name:
                        # run 自己没写字体时，真正印出来的字是段落/字符样式给的
                        style = run.style if run.style is not None else paragraph.style
                        name = style.font.name if style is not None else None
                except Exception:
                    name = None
            if name:
                counts[name] = counts.get(name, 0) + 1
    ordered = sorted(counts.items(), key=lambda kv: -kv[1])
    return [{"font": name, "runs": count} for name, count in ordered[:12]]


def _default_style_fonts(document):
    facts = []
    try:
        normal = document.styles["Normal"]
        size = normal.font.size
        facts.append({
            "style": "Normal",
            "font": normal.font.name,
            "sizePt": round(size.pt, 1) if size is not None else None,
        })
    except Exception:
        pass
    return facts


def _outline(document, limit=40):
    """按标题样式抽出大纲，给"这套模板的层级长什么样"用。"""
    entries = []
    for index, paragraph in enumerate(document.paragraphs):
        style = (paragraph.style.name or "") if paragraph.style is not None else ""
        text = paragraph.text.strip()
        if not text:
            continue
        if style.lower().startswith("heading") or style in ("Title", "Subtitle") or "标题" in style:
            entries.append({"paragraph": index, "style": style, "text": text[:120]})
        if len(entries) >= limit:
            break
    return entries


def _assets(document):
    images = 0
    try:
        images = len(document.inline_shapes)
    except Exception:
        images = 0
    return {
        "tables": len(document.tables),
        "images": images,
        "paragraphs": len(document.paragraphs),
        "sections": len(document.sections),
        "words": sum(len(p.text) for p in document.paragraphs),
    }


def inspect_docx(path, preview=12):
    """返回一份 docx 的版式事实（给模板反推与印前检查用）。"""
    try:
        document = Document(path)
    except Exception as error:
        raise EngineError(
            "unreadable-docx",
            "这个文件打不开，可能不是 .docx 或者已经损坏：" + str(path),
            "用 Word / WPS 打开它，另存为 .docx 再试") from error
    sections = [_section_facts(section) for section in document.sections]
    require(len(sections) > 0, "no-section", "这个文档没有可用的页面设置")

    previews = []
    for index, paragraph in enumerate(document.paragraphs):
        text = paragraph.text.strip()
        if not text:
            continue
        previews.append({"paragraph": index, "style": paragraph.style.name if paragraph.style else None, "text": text[:160]})
        if len(previews) >= preview:
            break

    return {
        "kind": "docx",
        "path": path,
        "sections": sections,
        "primary": sections[0],
        "fonts": _run_fonts(document),
        "defaultStyle": _default_style_fonts(document),
        "outline": _outline(document),
        "assets": _assets(document),
        "preview": previews,
        "metadata": metadata(document),
    }


def metadata(document):
    """文档属性：作者、单位、修订次数、最后修改人。净化功能靠它取证。"""
    props = document.core_properties
    return {
        "title": props.title or None,
        "author": props.author or None,
        "lastModifiedBy": props.last_modified_by or None,
        "revision": props.revision,
        "created": props.created.isoformat() if props.created else None,
        "modified": props.modified.isoformat() if props.modified else None,
        "comments": props.comments or None,
    }


def load(path):
    try:
        return Document(path)
    except Exception as error:  # 打不开就是打不开，给老师一句人话
        raise EngineError("docx-unreadable", f"打不开这个 Word 文档：{path}", str(error))

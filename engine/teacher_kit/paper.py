"""纸张与版心：中国教师认的是"8 开、16 开、A4、A3"，不是 Letter。

所有内部计算一律用毫米；落到 docx/pptx 时再转换成 EMU/twips。
"""

import re

from .errors import EngineError

# 名称 → (宽 mm, 高 mm)，一律按"纵向"记（宽 < 高），横向由 orientation 交换。
PAPERS = {
    "A4": (210.0, 297.0),
    "A3": (297.0, 420.0),
    "A5": (148.0, 210.0),
    "B5": (176.0, 250.0),
    "8K": (260.0, 368.0),   # 正度 8 开，试卷常用（横向→368×260）
    "16K": (184.0, 260.0),  # 正度 16 开
    "Letter": (215.9, 279.4),
}

# 名称 → (上, 下, 左, 右) mm。没写的用 DEFAULT_MARGINS。
MARGINS = {
    "A4": (25.4, 25.4, 31.8, 31.8),   # Word 中文版默认版心
    "A3": (25.4, 25.4, 31.8, 31.8),
    "8K": (24.0, 24.0, 27.0, 27.0),   # 试卷：上下 2.4cm、左右 2.7cm
    "16K": (24.0, 24.0, 27.0, 27.0),
    "B5": (25.4, 25.4, 25.4, 25.4),
    "A5": (20.0, 20.0, 20.0, 20.0),
    "Letter": (25.4, 25.4, 25.4, 25.4),
}

DEFAULT_PAPER = "A4"
DEFAULT_ORIENTATION = "portrait"
ORIENTATIONS = ("portrait", "landscape")


RE_MM = re.compile(r"^\s*(-?\d+(?:\.\d+)?)\s*(?:mm|毫米)?\s*$", re.IGNORECASE)
RE_CM = re.compile(r"^\s*(-?\d+(?:\.\d+)?)\s*(?:cm|厘米)\s*$", re.IGNORECASE)


def _margin_mm(value, label, default):
    """页边距读数字：老师可能写 24、"24mm"、"2.7cm"，也可能写错字。"""
    if value in (None, ""):
        return float(default)
    text = str(value).strip()
    match = RE_MM.match(text)
    if match:
        return float(match.group(1))
    match = RE_CM.match(text)
    if match:
        return float(match.group(1)) * 10.0
    raise EngineError(
        "bad-margin",
        "页边距只认数字（毫米）：" + label + " 写成了「" + text + "」",
        "比如 marginsMm: {top: 24, bottom: 24, left: 27, right: 27}")


def resolve(page):
    """page = {size?, orientation?, marginsMm?:{top,bottom,left,right}, columns?} → 完整版式参数。

    默认 A4 纵向：中国大陆教师新开一份文档就该是 A4，绝不能落到美式 Letter。
    """
    if isinstance(page, str):
        # 老板式写法：page: "8K"（只想换纸，不想写一整个对象）
        page = {"size": page}
    page = page or {}
    name = page.get("size") or DEFAULT_PAPER
    if name not in PAPERS:
        raise EngineError(
            "unknown-paper",
            "不认识的纸张大小：" + str(name),
            "可用：" + "、".join(PAPERS),
        )
    width, height = PAPERS[name]
    orientation = page.get("orientation") or DEFAULT_ORIENTATION
    if orientation not in ORIENTATIONS:
        raise EngineError("unknown-orientation", "不认识的方向：" + str(orientation), "可用：portrait、landscape")
    if orientation == "landscape":
        width, height = height, width

    defaults = MARGINS.get(name, (25.4, 25.4, 25.4, 25.4))
    given = page.get("marginsMm") or page.get("marginMm") or {}
    if not isinstance(given, dict):
        raise EngineError(
            "bad-margin",
            "页边距要写成对象：marginsMm: {top: 24, bottom: 24, left: 27, right: 27}",
            "四个方向单独给，单位毫米；不写就用这种纸的标配")
    top = _margin_mm(given.get("top"), "上", defaults[0])
    bottom = _margin_mm(given.get("bottom"), "下", defaults[1])
    left = _margin_mm(given.get("left"), "左", defaults[2])
    right = _margin_mm(given.get("right"), "右", defaults[3])

    for label, value in (("上", top), ("下", bottom), ("左", left), ("右", right)):
        if value < 0:
            raise EngineError("bad-margin", "页边距不能是负数：" + label + " " + str(value))

    if left + right >= width:
        raise EngineError(
            "margin-too-wide",
            "左右页边距加起来比纸还宽：" + str(round(left + right, 1)) + "mm ≥ " + str(width) + "mm",
            "把页边距调小，或换更大的纸",
        )
    if top + bottom >= height:
        raise EngineError(
            "margin-too-tall",
            "上下页边距加起来比纸还高：" + str(round(top + bottom, 1)) + "mm ≥ " + str(height) + "mm",
            "把页边距调小，或换更大的纸",
        )

    raw_columns = page.get("columns", 1)
    if raw_columns in (None, ""):
        raw_columns = 1
    try:
        columns = int(float(raw_columns))
    except (TypeError, ValueError):
        raise EngineError("bad-columns", "分栏数要填数字：写成了「" + str(raw_columns) + "」",
                          "试卷通常是 2 栏；要单栏就写 1 或者不写")
    if columns < 1 or columns > 4:
        raise EngineError("bad-columns", "分栏数只能是 1—4：" + str(columns), "试卷通常是 2 栏")

    return {
        "size": name,
        "orientation": orientation,
        "widthMm": round(width, 1),
        "heightMm": round(height, 1),
        "marginMm": {"top": top, "bottom": bottom, "left": left, "right": right},
        "columns": columns,
        "bodyWidthMm": round(width - left - right, 1),
    }

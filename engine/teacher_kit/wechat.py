"""微信公众号宣传稿：把老师给的素材排成一篇能直接粘进公众号后台的稿子。

边界（写死在实现里）
--------------------
* 不联网、不发布、不登录公众号：这里只产出文件，发不发、什么时候发由老师决定。
* 不编内容：正文是模型按老师给的材料写的，这个模块只负责结构、排版与体检。
  它不会把一句没影的话写漂亮，也不会替老师编活动照片说明。

产出
----
1. .html —— 全内联样式，宽度按公众号正文（677px）排好，浏览器打开"全选-复制-粘进编辑器"即可；
2. .docx —— 给要交给领导审、或者要存档的那份；
3. 体检报告 —— 标题长度、摘要、段落长度、绝对化用语、有没有落款，都是发出去之前最容易被挑的。

为什么公众号还要单独做：老师做宣传稿的真实卡点是"手机上看着挤成一片"和"标题被限流"，
不是"不会写字"。规则写在这里，比让模型每次自由发挥稳。
"""

import html as html_lib
import os
import re
import time

from .docx_build import build_docx
from .errors import EngineError, require
from .paths import default_output, guard_output_path

# 公众号正文区宽度（px）：粘进后台编辑器后不跑版就靠这个宽度
CONTENT_WIDTH = 677

THEMES = {
    "school": {"accent": "#1f5fa9", "soft": "#eef4fb", "ink": "#1f2328", "muted": "#6b7280"},
    "warm": {"accent": "#c2410c", "soft": "#fff1e7", "ink": "#2b2118", "muted": "#8b7355"},
    "green": {"accent": "#0f766e", "soft": "#e8f6f4", "ink": "#1c2b2a", "muted": "#5f7c7a"},
    "mono": {"accent": "#333333", "soft": "#f2f2f2", "ink": "#1a1a1a", "muted": "#777777"},
}

BLOCK_TYPES = ("lead", "heading", "text", "list", "quote", "note", "divider", "image")

# 教育宣传稿里最容易挨批的绝对化说法（广告法那套口径，发之前先自查一遍）
RE_ABSOLUTE = re.compile(
    r"(国家级|世界级|最高级|第一品牌|唯一|独家|顶级|绝无仅有|史无前例|永久|彻底解决|包过|保过|100%|百分百|万能"
    r"|最好|最佳|最棒|最强|最优|绝对|保证|零风险|根治|领先|全国第一|全市第一|第一名)"
)


def _text(value):
    return str(value).strip() if value not in (None, "") else ""


def _normalize_blocks(blocks):
    """接受紧凑写法：字符串当一段正文，字典缺 type 就当 text。"""
    out = []
    warnings = []
    for index, raw in enumerate(blocks or []):
        if isinstance(raw, str):
            text = raw.strip()
            if text:
                out.append({"type": "text", "text": text})
            continue
        if not isinstance(raw, dict):
            warnings.append({"code": "block-dropped", "level": "warn",
                             "message": "第 " + str(index + 1) + " 段不是我认得的写法，已经跳过"})
            continue
        kind = str(raw.get("type") or "text").strip().lower()
        if kind not in BLOCK_TYPES:
            warnings.append({"code": "block-type-unknown", "level": "info",
                             "message": "第 " + str(index + 1) + " 段的类型「" + kind + "」不认识，按正文排了"})
            kind = "text"
        block = {"type": kind}
        if kind == "divider":
            out.append(block)
            continue
        if kind == "image":
            path = _text(raw.get("path") or raw.get("src"))
            caption = _text(raw.get("caption")) or _text(raw.get("text"))
            alt = _text(raw.get("alt"))
            if not path and not caption:
                warnings.append({"code": "image-empty", "level": "warn",
                                 "message": "第 " + str(index + 1) + " 段是配图，可既没给图片文件也没写说明，已经跳过"})
                continue
            if path:
                block["path"] = path
            if caption:
                block["caption"] = caption
            if alt:
                block["alt"] = alt
            if raw.get("widthMm") not in (None, ""):
                block["widthMm"] = raw["widthMm"]
            out.append(block)
            continue
        if kind == "list":
            items = raw.get("items")
            if isinstance(items, str):
                items = [line for line in items.splitlines() if line.strip()]
            items = [_text(item) for item in (items or []) if _text(item)]
            if not items:
                text = _text(raw.get("text"))
                items = [line.strip() for line in text.splitlines() if line.strip()]
            if not items:
                continue
            block["items"] = items
            if raw.get("ordered"):
                block["ordered"] = True
        else:
            text = _text(raw.get("text"))
            if not text:
                continue
            block["text"] = text
        out.append(block)
    return out, warnings


def _paragraph_report(blocks, digest, title):
    """体检：发出去之前最容易被挑的几件事，都在这里查。"""
    warnings = []
    title_len = len(title)
    if not title:
        warnings.append({"code": "title-missing", "level": "error",
                         "message": "还没有标题——公众号标题最多 64 字，建议 12—22 字"})
    elif title_len > 40:
        warnings.append({"code": "title-too-long", "level": "error",
                         "message": "标题 " + str(title_len) + " 字太长了，手机上一行放不下，压到 22 字以内点击率才不掉"})
    elif title_len > 24:
        warnings.append({"code": "title-too-long", "level": "warn",
                         "message": "标题 " + str(title_len) + " 字偏长（手机列表页大概只显示 20 字）"})
    if not digest:
        warnings.append({"code": "digest-missing", "level": "warn",
                         "message": "没写摘要：分享到群里时只显示标题，摘要能多带一句话"})
    elif len(digest) > 120:
        warnings.append({"code": "digest-too-long", "level": "warn",
                         "message": "摘要 " + str(len(digest)) + " 字，公众号只显示前 120 字"})
    body_paras = [block for block in blocks if block["type"] in ("text", "lead", "quote")]
    if not body_paras:
        warnings.append({"code": "body-empty", "level": "error",
                         "message": "正文是空的：至少给一段话说清这次活动/通知是什么"})
    longest = 0
    for block in body_paras:
        length = len(block.get("text") or "")
        longest = max(longest, length)
        if length > 250:
            warnings.append({"code": "paragraph-too-long", "level": "warn",
                             "message": "有一段 " + str(length) + " 字，手机上就是一大坨，拆成两段每段 3—4 行更好读"})
    for block in blocks:
        if block["type"] != "list":
            continue
        for item in block.get("items") or []:
            if len(item) > 40:
                warnings.append({"code": "list-item-too-long", "level": "info",
                                 "message": "要点偏长（" + str(len(item)) + " 字）：" + item[:16] + "…，手机上会折两行"})
    whole = " ".join([title, digest] + [block.get("text") or "" for block in blocks]
                     + [" ".join(block.get("items") or []) for block in blocks])
    hits = sorted(set(RE_ABSOLUTE.findall(whole)))
    if hits:
        warnings.append({"code": "absolute-claim", "level": "warn",
                         "message": "有绝对化说法：" + "、".join(hits) + "（宣传稿里容易被挑，换成具体数字更稳）",
                         "words": hits})
    if not any(block["type"] == "note" for block in blocks):
        warnings.append({"code": "no-signoff", "level": "info",
                         "message": "结尾没有「文/图 某某」「供稿：某某」这类落款行，发出去显得没头没尾"})
    total = sum(len(block.get("text") or "") for block in body_paras)
    total += sum(len(item) for block in blocks if block["type"] == "list"
                 for item in block.get("items") or [])
    stats = {
        "images": sum(1 for block in blocks if block["type"] == "image"),
        "titleChars": title_len,
        "digestChars": len(digest),
        "paragraphs": len(body_paras),
        "longestParagraph": longest,
        "bodyChars": total,
        "readSeconds": int(round(total / 300.0 * 60)) if total else 0,
    }
    return stats, warnings


def _theme(name):
    key = _text(name).lower() or "school"
    return THEMES.get(key) or THEMES["school"]


def _p(style, text):
    return '<p style="' + style + '">' + html_lib.escape(text) + "</p>"


def render_html(post):
    """全内联样式的 HTML：公众号编辑器只吃内联样式，外链 CSS 粘过去就丢。"""
    theme = _theme(post.get("theme"))
    accent, soft, ink, muted = theme["accent"], theme["soft"], theme["ink"], theme["muted"]
    esc = html_lib.escape
    parts = ['<!DOCTYPE html>',
             '<html lang="zh-CN"><head><meta charset="utf-8">',
             '<meta name="viewport" content="width=device-width, initial-scale=1">',
             "<title>" + esc(post.get("title") or "") + "</title></head>",
             '<body style="margin:0;padding:24px 12px;background:#f5f6f7;">']
    box_style = ("max-width:" + str(CONTENT_WIDTH) + "px;margin:0 auto;background:#ffffff;"
                 "padding:32px 22px;box-sizing:border-box;font-size:16px;line-height:1.78;"
                 "font-family:-apple-system,BlinkMacSystemFont,'PingFang SC','Microsoft YaHei',sans-serif;"
                 "color:" + ink + ";")
    parts.append('<section style="' + box_style + '">')
    parts.append(_p("margin:0 0 6px;font-size:19px;font-weight:700;letter-spacing:0.2px;",
                    post.get("title") or ""))
    if post.get("digest"):
        parts.append(_p("margin:0 0 18px;font-size:13px;color:" + muted + ";", post["digest"]))
    parts.append('<section style="height:3px;width:44px;background:' + accent + ';margin:0 0 22px;"></section>')
    for block in post.get("blocks") or []:
        kind = block["type"]
        if kind == "divider":
            parts.append('<p style="margin:22px 0;border-top:1px dashed ' + muted + ';"></p>')
            continue
        if kind == "heading":
            parts.append(_p("margin:22px 0 10px;font-size:17px;font-weight:700;color:" + accent + ";",
                            block["text"]))
            continue
        if kind == "image":
            # 本地图片粘不进公众号编辑器（它不认本地路径，只认自己上传的素材），
            # 所以这里留一个写得清楚的位置标记，老师在后台一处点「图片」就能补上。
            label = block.get("caption") or block.get("alt") or "活动照片"
            parts.append('<section style="margin:20px 0;padding:24px 14px;border:1px dashed ' + muted + ";"
                         'background:' + soft + ';text-align:center;font-size:13px;color:' + muted + ';">'
                         + esc("【配图】" + label)
                         + '<br><span style="font-size:11px;">（公众号要在这里点「图片」上传，本地图片粘不进去）</span>'
                         "</section>")
            continue
        if kind == "lead":
            inner = _p("margin:0;font-size:15px;color:" + ink + ";", block["text"])
            wrap = "border-left:3px solid " + accent + ";background:" + soft + ";padding:12px 14px;margin:0 0 18px;"
            parts.append('<section style="' + wrap + '">' + inner + "</section>")
            continue
        if kind == "quote":
            inner = _p("margin:0;font-size:15px;color:" + muted + ";", block["text"])
            wrap = "border-left:3px solid " + muted + ";padding:6px 14px;margin:18px 0;"
            parts.append('<section style="' + wrap + '">' + inner + "</section>")
            continue
        if kind == "note":
            parts.append(_p("margin:20px 0 0;font-size:13px;color:" + muted + ";", block["text"]))
            continue
        if kind == "list":
            rows = []
            for index, item in enumerate(block.get("items") or []):
                bullet = (str(index + 1) + ".") if block.get("ordered") else "•"
                mark = '<span style="color:' + accent + ';font-weight:700;">' + bullet + "</span> "
                rows.append('<p style="margin:0 0 8px;padding-left:18px;text-indent:-18px;">'
                            + mark + esc(item) + "</p>")
            parts.append('<section style="margin:14px 0 18px;font-size:15.5px;">' + "".join(rows) + "</section>")
            continue
        parts.append(_p("margin:0 0 17px;font-size:16px;letter-spacing:0.3px;", block["text"]))
    footer = []
    if post.get("author"):
        footer.append("文 / " + post["author"])
    if post.get("signature"):
        footer.append(post["signature"])
    if footer:
        body = "<br>".join(esc(item) for item in footer)
        wrap = "margin:26px 0 0;padding-top:14px;border-top:1px solid #eceef0;font-size:13px;color:" + muted + ";"
        parts.append('<p style="' + wrap + '">' + body + "</p>")
    parts.append("</section></body></html>")
    return "\n".join(parts)


def _docx_blocks(post):
    """同一份内容排成 Word：给要审稿、要存档的那份。"""
    blocks = [{"type": "paragraph", "align": "center",
               "runs": [{"text": post.get("title") or "", "bold": True, "sizePt": 16, "font": "宋体"}]}]
    if post.get("digest"):
        blocks.append({"type": "paragraph", "align": "center", "spaceAfterPt": 10,
                       "runs": [{"text": post["digest"], "sizePt": 10.5, "font": "宋体"}]})
    for block in post.get("blocks") or []:
        kind = block["type"]
        if kind == "divider":
            blocks.append({"type": "paragraph", "align": "center",
                           "runs": [{"text": "—— —— ——", "sizePt": 10, "font": "宋体"}]})
            continue
        if kind == "heading":
            blocks.append({"type": "paragraph", "spaceAfterPt": 4,
                           "runs": [{"text": block["text"], "bold": True, "sizePt": 13, "font": "宋体"}]})
            continue
        if kind == "image":
            path = block.get("path")
            caption = block.get("caption") or block.get("alt") or ""
            if path and os.path.isfile(str(path)):
                item = {"type": "picture", "path": path, "caption": caption}
                if block.get("widthMm") not in (None, ""):
                    item["widthMm"] = block["widthMm"]
                blocks.append(item)
            else:
                blocks.append({"type": "paragraph", "align": "center", "spaceAfterPt": 6,
                               "runs": [{"text": "【配图】" + (caption or "活动照片"), "sizePt": 10.5,
                                         "font": "宋体"}]})
            continue
        if kind == "list":
            for index, item in enumerate(block.get("items") or []):
                mark = (str(index + 1) + ".") if block.get("ordered") else "·"
                blocks.append({"type": "paragraph", "lineSpacingPt": 22,
                               "runs": [{"text": mark + " " + item, "sizePt": 12, "font": "宋体"}]})
            continue
        size = 12 if kind == "text" else 10.5
        blocks.append({"type": "paragraph", "lineSpacingPt": 22 if kind == "text" else 20,
                       "runs": [{"text": block["text"], "sizePt": size, "font": "宋体"}]})
    footer = []
    if post.get("author"):
        footer.append("文 / " + post["author"])
    if post.get("signature"):
        footer.append(post["signature"])
    if footer:
        blocks.append({"type": "paragraph", "align": "right", "spaceAfterPt": 0,
                       "runs": [{"text": "\n".join(footer), "sizePt": 10.5, "font": "宋体"}]})
    return blocks


def wechat_check(payload):
    """只体检不产出：老师贴一段自己写的稿子，先看哪里会被手机吃掉。"""
    text = _text(payload.get("text"))
    title = _text(payload.get("title"))
    blocks = payload.get("blocks")
    if not blocks and text:
        # 贴进来的整段稿子：空行分段，第一行当标题（老师习惯这么贴）
        lines = [line.strip() for line in text.splitlines()]
        if not title and lines:
            title = lines[0]
            lines = lines[1:]
        blocks = []
        buffer = []
        for line in lines:
            if not line:
                if buffer:
                    blocks.append(" ".join(buffer))
                    buffer = []
                continue
            buffer.append(line)
        if buffer:
            blocks.append(" ".join(buffer))
    blocks, warnings = _normalize_blocks(blocks)
    stats, more = _paragraph_report(blocks, _text(payload.get("digest")), title)
    return {"title": title, "digest": _text(payload.get("digest")), "blocks": blocks,
            "stats": stats, "warnings": warnings + more}


def _check_pictures(blocks):
    """出成品之前把配图挨个真正打开一次。

    以前图片坏了要等到 python-docx 里才炸，回执把二进制的解码栈甩给老师
    （「'utf-8' codec can't decode byte 0xae」），而且 html 已经写了一半。
    这里先报清楚「这张不是图片 / 已经损坏」，调用方也就不会留下半成品。
    """
    broken = []
    try:
        from PIL import Image
    except ImportError:      # 没有 Pillow 就交给 python-docx 兜底，别把出稿卡死
        return broken
    for block in blocks:
        if block.get("type") != "image":
            continue
        path = str(block.get("path") or "")
        if not path or not os.path.isfile(path):
            continue
        try:
            with Image.open(path) as image:
                image.verify()
        except Exception:
            broken.append(path)
    if broken:
        raise EngineError(
            "picture-unreadable",
            "配图打不开：" + "、".join(os.path.basename(path) for path in broken),
            "这张多半不是图片文件、或者已经损坏了：用照片查看器打开确认一下，"
            "另存成 .jpg / .png 再选一次")
    return broken


def wechat_render(payload):
    """出一份公众号稿：HTML（直接粘）+ docx（审稿存档）。"""
    title = _text(payload.get("title")) or "公众号稿件"
    digest = _text(payload.get("digest"))
    blocks, block_warnings = _normalize_blocks(payload.get("blocks"))
    require(bool(blocks), "empty-post", "还没有正文：至少给一段话说清这次要宣传什么")
    stats, warnings = _paragraph_report(blocks, digest, title)
    warnings = block_warnings + warnings
    post = {
        "title": title,
        "digest": digest,
        "author": _text(payload.get("author")),
        "signature": _text(payload.get("signature")),
        "theme": _text((payload.get("options") or {}).get("theme")) or "school",
        "blocks": blocks,
    }
    formats = payload.get("formats") or ["html", "docx"]
    if isinstance(formats, str):
        formats = [formats]
    formats = [str(item).lower() for item in formats if str(item).lower() in ("html", "docx")]
    require(bool(formats), "no-format", "要哪种成品？html（粘公众号）或 docx（审稿存档）")

    base = payload.get("outputPath")
    if base:
        base = os.path.abspath(base)
        stem = os.path.splitext(base)[0]
    else:
        extension = ".docx" if "docx" in formats else ".html"
        base = default_output("公众号", title, extension)
        stem = os.path.splitext(base)[0]
    guard_output_path(stem + ".html" if "html" in formats else stem + ".docx")
    folder = os.path.dirname(stem)
    if folder and not os.path.isdir(folder):
        os.makedirs(folder, exist_ok=True)
    for block in blocks:
        if block["type"] != "image":
            continue
        path = block.get("path")
        if path and not os.path.isfile(str(path)):
            warnings.append({"code": "image-missing", "level": "warn",
                             "message": "配图找不到：" + os.path.basename(str(path)) + "（稿子里留了位置，图要重新选一次）"})
    _check_pictures(blocks)

    outputs = {}
    overwrite = bool(payload.get("overwrite", False))
    html_path = stem + ".html" if "html" in formats else None
    docx_path = stem + ".docx" if "docx" in formats else None
    for target in (html_path, docx_path):
        if target and os.path.exists(target) and not overwrite:
            raise EngineError(
                "output-exists",
                "这个位置已经有同名稿子了，我不动它：" + os.path.basename(target),
                "换个标题（会自动换个文件名），或者加 overwrite:true 明确允许覆盖")
    # 两份成品都先写临时名、都成功再改名：以前 html 先落盘、docx 因为图片坏掉抛错，
    # 老师手里留下半份稿子，重跑还被 output-exists 挡住。坏图在写之前就拦住。
    stamp = str(os.getpid()) + "-" + str(time.time_ns())
    temp_html = stem + ".tk-" + stamp + ".html" if html_path else None
    temp_docx = stem + ".tk-" + stamp + ".docx" if docx_path else None
    written = []
    try:
        if temp_html:
            with open(temp_html, "w", encoding="utf-8") as handle:
                handle.write(render_html(post))
            written.append((temp_html, html_path))
        if temp_docx:
            built = build_docx({
                "outputPath": temp_docx,
                "page": {"size": "A4", "orientation": "portrait", "columns": 1},
                "blocks": _docx_blocks(post),
                "properties": {"title": title, "author": post["author"] or "教师助手"},
                "defaultFooter": False,  # 这是往公众号里贴的稿子，不是试卷，别印「出卷老师」
                "overwrite": True,
            })
            outputs["docxBytes"] = built["bytes"]
            written.append((temp_docx, docx_path))
        for temp, final in written:
            os.replace(temp, final)
            if os.path.exists(temp):
                try:
                    os.remove(temp)
                except OSError:
                    pass
    except BaseException:
        for temp, _final in written:
            try:
                if os.path.exists(temp):
                    os.remove(temp)
            except OSError:
                pass
        raise
    if html_path:
        outputs["htmlPath"] = html_path
        outputs["htmlBytes"] = os.path.getsize(html_path)
    if docx_path:
        outputs["outputPath"] = docx_path
    result = {
        "title": title,
        "digest": digest,
        "blocks": blocks,
        "stats": stats,
        "warnings": warnings,
        "theme": post["theme"],
        "formats": formats,
    }
    result.update(outputs)
    return result

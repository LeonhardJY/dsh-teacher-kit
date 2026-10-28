"""课件（HTML 形态）：借来的电脑没装 Office、要发班级群时，双击就能放。

为什么要有这一份：pptx 是正式交付形态，但它要求对方装了 PowerPoint / WPS。
公开课教室的电脑常常只有浏览器，班群里发 pptx 家长也打不开，
所以同一份分页结果要能再落成一个 HTML 文件：不联网、不引外部 CSS/JS、双击即放映。

分页逻辑一律复用 lesson.plan_deck —— 这里只管"怎么显示"，绝不重排页数：
同一段话在 pptx 和 html 里必须是一样多页，老师拿两份对得上。

版式取舍：
- 以 1280×720 为基准排版，再用 transform: scale() 整体缩放。字号写死 px，
  换窗口只缩不重排，老师在哪台机器上看到的换行位置都一致（投影不跳版）。
- 打印按 297mm × 167mm：A4 横放宽度 297mm，297 / 16 * 9 ≈ 167mm，正好 16:9，
  一页一张、不裁字，也不用浏览器再缩一次。
- 浅底深字：投影对比度够，打印省墨（教室的墨盒是老师自己买的）。
"""

import os

from html import escape

from .errors import EngineError, require
from .lesson import plan_deck
from .paths import default_output

# 放映基准尺寸：1280×720 是投影/录屏最通用的 16:9 栅格，缩放只作用在这一层。
BASE_WIDTH = 1280
BASE_HEIGHT = 720
MAX_BULLET_LEVEL = 2
# 超过这个条数就整体收紧字号：宁可小一号，也别让最后两条掉出屏幕
DENSE_BULLET_COUNT = 7

CSS = """@charset "utf-8";
:root {
  --ink: #16202c;
  --ink-soft: #4a5a68;
  --accent: #1b5fae;
  --paper: #ffffff;
  --line: #dbe3ec;
  --scale: 1;
}
* { box-sizing: border-box; }
html, body { margin: 0; padding: 0; }
html { -webkit-text-size-adjust: 100%; }
body {
  background: #eef1f5;
  color: var(--ink);
  font-family: system-ui, -apple-system, "Segoe UI", "Microsoft YaHei", "PingFang SC", sans-serif;
  overflow: hidden;
}
.deck { position: fixed; inset: 0; }

/* 一页 = 1280×720 的世界，整体缩放；display:none 的页不参与布局，省内存 */
.slide {
  position: absolute;
  left: 50%;
  top: 50%;
  width: 1280px;
  height: 720px;
  display: none;
  flex-direction: column;
  padding: 56px 76px 64px;
  background: var(--paper);
  border-radius: 10px;
  box-shadow: 0 18px 60px rgba(16, 34, 54, 0.18);
  overflow: hidden;
  transform: translate(-50%, -50%) scale(var(--scale, 1));
  transform-origin: center center;
}
.slide.active { display: flex; }
.slide.cover { justify-content: center; align-items: center; text-align: center; }

.slide-head { margin: 0 0 26px; padding-bottom: 14px; border-bottom: 3px solid var(--line); }
.slide.cover .slide-head { border-bottom: 0; margin-bottom: 10px; padding-bottom: 0; }
.slide-title {
  margin: 0;
  font-size: 52px;
  line-height: 1.22;
  font-weight: 700;
  letter-spacing: 0.5px;
}
.slide.cover .slide-title { font-size: 76px; line-height: 1.16; letter-spacing: 2px; }
.slide-subtitle {
  margin: 14px 0 0;
  font-size: 30px;
  font-weight: 400;
  color: var(--ink-soft);
}
.slide.cover .slide-subtitle { font-size: 32px; letter-spacing: 1px; }

.slide-body { flex: 1; min-height: 0; display: flex; flex-direction: column; justify-content: center; }
ul.bullets { list-style: none; margin: 0; padding: 0; display: flex; flex-direction: column; gap: 20px; }
li.bullet { position: relative; padding-left: 42px; font-size: 34px; line-height: 1.45; }
li.bullet::before {
  content: "";
  position: absolute;
  left: 10px;
  top: 0.62em;
  width: 14px;
  height: 14px;
  border-radius: 50%;
  background: var(--accent);
}
li.bullet.level-1 { margin-left: 46px; font-size: 28px; color: var(--ink-soft); }
li.bullet.level-1::before { width: 10px; height: 10px; border-radius: 2px; background: #7b8c9d; }
li.bullet.level-2 { margin-left: 92px; font-size: 24px; color: var(--ink-soft); }
li.bullet.level-2::before {
  width: 10px; height: 10px; border-radius: 50%;
  background: transparent; border: 2px solid #9aa9b8;
}
/* 要点多时整体收紧：一页塞到 10 条以上还要投出去，字号大就看不全 */
.slide.dense li.bullet { font-size: 26px; }
.slide.dense li.bullet.level-1 { font-size: 22px; }
.slide.dense ul.bullets { gap: 12px; }

.pageno {
  position: absolute;
  right: 36px;
  bottom: 22px;
  font-size: 20px;
  color: #8b9bab;
  letter-spacing: 1px;
  font-variant-numeric: tabular-nums;
}

/* 讲稿默认不显示：投影时学生不该看到老师的小抄，按 N 才出来提词 */
.notes {
  display: none;
  position: absolute;
  left: 40px;
  right: 40px;
  bottom: 54px;
  max-height: 46%;
  overflow: auto;
  padding: 18px 24px;
  border-radius: 10px;
  background: rgba(18, 28, 42, 0.93);
  color: #f2f6fa;
  font-size: 23px;
  line-height: 1.55;
  text-align: left;
}
body.show-notes .notes { display: block; }
.notes-label { display: block; font-size: 15px; letter-spacing: 3px; color: #9db3c9; margin-bottom: 6px; }
.notes p { margin: 0; }

.controls {
  position: fixed;
  right: 20px;
  bottom: 16px;
  z-index: 20;
  display: flex;
  align-items: center;
  gap: 8px;
  padding: 6px 8px;
  border: 1px solid var(--line);
  border-radius: 999px;
  background: rgba(255, 255, 255, 0.94);
  box-shadow: 0 6px 18px rgba(16, 34, 54, 0.12);
}
.controls button {
  width: 40px;
  height: 40px;
  border: 0;
  border-radius: 50%;
  background: #eaf0f7;
  color: var(--accent);
  font-size: 20px;
  line-height: 1;
  cursor: pointer;
}
.controls button:hover { background: #dbe7f4; }
.counter { padding: 0 8px; font-size: 16px; color: var(--ink-soft); font-variant-numeric: tabular-nums; }

.hint {
  position: fixed;
  left: 18px;
  bottom: 18px;
  z-index: 20;
  margin: 0;
  font-size: 14px;
  color: #7d8c9b;
}

@media print {
  /* 297mm × 167mm = A4 横放，且正好 16:9：一页一张，不裁字 */
  @page { size: 297mm 167mm; margin: 0; }
  html, body { background: #fff; overflow: visible; }
  body { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
  .deck { position: static; }
  .controls, .hint, .pageno, .counter, .notes { display: none !important; }
  .slide {
    display: flex !important;
    position: static;
    width: auto;
    height: 167mm;
    padding: 12mm 18mm;
    border-radius: 0;
    box-shadow: none;
    transform: none !important;
    page-break-after: always;
    break-after: page;
  }
  /* 最后一页不再断页，否则打印机会多吐一张空白纸 */
  .slide:last-of-type { page-break-after: auto; break-after: auto; }
  .slide-head { margin-bottom: 6mm; padding-bottom: 3mm; }
  .slide-title { font-size: 26pt; }
  .slide.cover .slide-title { font-size: 36pt; }
  .slide-subtitle { font-size: 15pt; margin-top: 4mm; }
  .slide.cover .slide-subtitle { font-size: 15pt; }
  ul.bullets { gap: 4mm; }
  li.bullet { font-size: 16pt; padding-left: 8mm; }
  li.bullet::before { left: 2mm; top: 0.6em; width: 3mm; height: 3mm; }
  li.bullet.level-1 { font-size: 14pt; margin-left: 8mm; }
  li.bullet.level-2 { font-size: 12pt; margin-left: 16mm; }
  .slide.dense li.bullet { font-size: 13pt; }
}
"""

SCRIPT = """(function () {
  "use strict";
  var slides = Array.prototype.slice.call(document.querySelectorAll(".slide"));
  if (!slides.length) { return; }
  var counter = document.getElementById("counter");
  var current = 0;

  // 缩放而不是改字号：版式按 1280×720 排一次，投影/笔记本只缩不重排
  function fit() {
    var width = window.innerWidth || 1280;
    var height = window.innerHeight || 720;
    var scale = Math.min(width / 1280, height / 720) * 0.98;
    if (!isFinite(scale) || scale <= 0) { scale = 1; }
    document.documentElement.style.setProperty("--scale", String(scale));
  }

  function show(index) {
    if (index < 0) { index = 0; }
    if (index > slides.length - 1) { index = slides.length - 1; }
    for (var i = 0; i < slides.length; i++) {
      slides[i].classList.toggle("active", i === index);
    }
    current = index;
    if (counter) { counter.textContent = (index + 1) + " / " + slides.length; }
    var want = "#" + (index + 1);
    // replaceState：翻页不往历史里塞 N 条，按后退键能离开这份课件
    if (window.location.hash !== want) {
      try { window.history.replaceState(null, "", want); }
      catch (error) { window.location.hash = want; }
    }
  }

  function hashIndex() {
    var match = /^#([0-9]+)$/.exec(window.location.hash || "");
    return match ? parseInt(match[1], 10) - 1 : -1;
  }

  function step(delta) { show(current + delta); }

  document.addEventListener("keydown", function (event) {
    if (event.altKey || event.ctrlKey || event.metaKey) { return; }
    var key = event.key;
    if (key === "ArrowRight" || key === "ArrowDown" || key === "PageDown" ||
        key === " " || key === "Spacebar" || key === "Enter") {
      event.preventDefault();
      step(1);
      return;
    }
    if (key === "ArrowLeft" || key === "ArrowUp" || key === "PageUp" || key === "Backspace") {
      event.preventDefault();
      step(-1);
      return;
    }
    if (key === "Home") { event.preventDefault(); show(0); return; }
    if (key === "End") { event.preventDefault(); show(slides.length - 1); return; }
    if (key === "n" || key === "N") {
      event.preventDefault();
      document.body.classList.toggle("show-notes");
    }
  });

  var prev = document.getElementById("prev");
  var next = document.getElementById("next");
  if (prev) { prev.addEventListener("click", function () { step(-1); }); }
  if (next) { next.addEventListener("click", function () { step(1); }); }

  window.addEventListener("resize", fit);
  window.addEventListener("orientationchange", fit);
  // 手改地址栏的 #7、或浏览器前进后退，都能跟上
  window.addEventListener("hashchange", function () {
    var index = hashIndex();
    if (index >= 0) { show(index); }
  });

  fit();
  var start = hashIndex();
  show(start >= 0 ? start : 0);
})();
"""


def _esc(value):
    """老师写的每个字都当文本：<script>、&、引号必须原样显示，绝不能变成标签。"""
    return escape("" if value is None else str(value))


def _normalize_bullets(raw):
    """把要点统一成 [(level, text)]。

    plan_deck 给的是 {text, level}；外面直接调 render_html 时可能只给字符串，
    这里一并接住，免得调用方为了排版细节再包一层。
    """
    items = []
    for entry in raw or []:
        if isinstance(entry, dict):
            text = entry.get("text")
            level = entry.get("level") or 0
        else:
            text = entry
            level = 0
        text = "" if text is None else str(text).strip()
        if not text:
            continue
        try:
            level = int(level)
        except (TypeError, ValueError):
            level = 0
        # 层级太深在投影上就是往右缩到没有地方写，收在 2 级
        items.append((min(max(level, 0), MAX_BULLET_LEVEL), text))
    return items


def _slide_html(slide, index, total):
    """一张 <section>：标题区 + 要点区 + 页码 + 讲稿。"""
    if not isinstance(slide, dict):
        slide = {"title": slide}
    is_cover = str(slide.get("layout") or "") == "title"
    bullets = _normalize_bullets(slide.get("bullets"))
    classes = ["slide"] + (["cover"] if is_cover else [])
    if len(bullets) > DENSE_BULLET_COUNT:
        classes.append("dense")

    title = slide.get("title") or ("课件" if is_cover else "")
    subtitle = slide.get("subtitle")
    notes = slide.get("notes")
    notes = "" if notes is None else str(notes).strip()

    heading = "h1" if is_cover else "h2"
    parts = [
        '<section class="' + " ".join(classes) + '"'
        ' id="slide-' + str(index) + '" data-index="' + str(index) + '"'
        ' aria-label="第 ' + str(index) + " 页，共 " + str(total) + ' 页">',
        '<header class="slide-head">',
        "<" + heading + ' class="slide-title">' + _esc(title) + "</" + heading + ">",
    ]
    if subtitle:
        parts.append('<p class="slide-subtitle">' + _esc(subtitle) + "</p>")
    parts.append("</header>")

    parts.append('<div class="slide-body">')
    if bullets:
        parts.append('<ul class="bullets">')
        for level, text in bullets:
            parts.append('<li class="bullet level-' + str(level) + '">' + _esc(text) + "</li>")
        parts.append("</ul>")
    parts.append("</div>")

    parts.append('<div class="pageno">' + str(index) + " / " + str(total) + "</div>")
    if notes:
        parts.append('<aside class="notes">')
        parts.append('<span class="notes-label">讲稿</span>')
        parts.append("<p>" + _esc(notes) + "</p>")
        parts.append("</aside>")
    parts.append("</section>")
    return "\n".join(parts)


def render_html(slides, title=None, options=None):
    """slides（lesson.plan_deck 的产出）→ 一份完整、自包含的单文件 HTML 字符串。

    不落盘、不改 slides：纯函数，测试与宿主都能直接拿到字符串自己处理。
    """
    options = options or {}
    require(isinstance(slides, (list, tuple)) and len(slides) > 0,
            "slides-required",
            "没有可放映的页（内容还没排成页）",
            "先调用 lesson.plan_deck 把内容分页，再把 slides 交进来")

    deck_title = title or options.get("title") or "课件"
    slides = list(slides)

    # 和 pptx 保持一致：options.notes 落到封面备注（开场那句"这节课 40 分钟"）
    cover_notes = options.get("notes")
    if cover_notes:
        slides[0] = dict(slides[0]) if isinstance(slides[0], dict) else {"title": slides[0]}
        slides[0]["notes"] = str(cover_notes)

    total = len(slides)
    body = [
        '<main class="deck" id="deck">',
        "\n".join(_slide_html(slide, index, total) for index, slide in enumerate(slides, start=1)),
        "</main>",
        '<nav class="controls" aria-label="翻页">',
        '<button type="button" id="prev" aria-label="上一页">←</button>',
        '<span class="counter" id="counter">1 / ' + str(total) + "</span>",
        '<button type="button" id="next" aria-label="下一页">→</button>',
        "</nav>",
        '<p class="hint" id="hint">← → / 空格 翻页 · Home / End 首末页 · N 显示讲稿</p>',
    ]

    return (
        "<!DOCTYPE html>\n"
        '<html lang="zh-CN">\n'
        "<head>\n"
        '<meta charset="utf-8">\n'
        '<meta name="viewport" content="width=device-width, initial-scale=1">\n'
        '<meta name="generator" content="teacher-kit">\n'
        "<title>" + _esc(deck_title) + "</title>\n"
        "<style>\n" + CSS + "</style>\n"
        "</head>\n"
        '<body class="deck-body">\n'
        + "\n".join(body) + "\n"
        "<script>\n" + SCRIPT + "</script>\n"
        "</body>\n"
        "</html>\n"
    )


def _ensure_parent(path):
    parent = os.path.dirname(os.path.abspath(path))
    if parent and not os.path.isdir(parent):
        os.makedirs(parent, exist_ok=True)


def build_deck_html(payload):
    """payload = {text, title?, subtitle?, options?, outputPath?, overwrite?}

    → {outputPath, slides, bytes, warnings, size}
    空 text 不在这里拦：plan_deck 抛的 text-required 带着"把内容粘进来"的提示，
    吞掉再自己抛一遍只会把提示丢掉。
    """
    payload = payload or {}
    options = payload.get("options") or {}
    planned = plan_deck(payload)
    title = payload.get("title") or options.get("title") or "课件"

    html_text = render_html(planned["slides"], title=title, options=options)

    output = payload.get("outputPath") or default_output("课件", title, ".html")
    _ensure_parent(output)
    if os.path.exists(output) and not payload.get("overwrite", False):
        raise EngineError("output-exists", "目标文件已存在，不覆盖：" + str(output),
                          "换一个文件名，或显式允许覆盖")

    # newline="\n"：同一份课件在 Windows / Mac 上字节数一致，发群里不会一个 12KB 一个 13KB
    with open(output, "w", encoding="utf-8", newline="\n") as handle:
        handle.write(html_text)

    return {
        "outputPath": os.path.abspath(output),
        "slides": len(planned["slides"]),
        "bytes": os.path.getsize(output),
        "warnings": list(planned["warnings"]),
        "size": "16:9",
    }

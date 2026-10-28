#!/usr/bin/env python
"""课件 HTML：跑 python test_lesson_html.py。

全是"读文件断言"：不启浏览器，只用字符串查结构/CSS/脚本，
老师在借来的电脑上双击能放，靠的就是这份单文件里什么都不缺。
"""

import os
import re
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
ENGINE = os.path.dirname(HERE)
sys.path.insert(0, ENGINE)

# 先 isolate 文档目录：测试绝不能往老师真实的「文档/教师助手」里丢文件
WORK = tempfile.mkdtemp(prefix="lesson-html-")
os.environ["TEACHER_KIT_HOME"] = os.path.join(WORK, "教师助手")

from teacher_kit.errors import EngineError  # noqa: E402
from teacher_kit.lesson import plan_deck  # noqa: E402
from teacher_kit.lesson_html import build_deck_html, render_html  # noqa: E402

FAILS = []


def check(condition, label, detail=""):
    if condition:
        print("  ✓ " + label)
    else:
        FAILS.append(label)
        print("  ✗ " + label + (" — " + str(detail) if detail else ""))


def expect_error(fn, code, label):
    try:
        fn()
    except EngineError as error:
        check(error.code == code, label, "拿到的是 " + error.code)
        return error
    check(False, label, "没有抛错")
    return None


# 封面 + 四个小节 = 5 页；"讲："那行落进备注
FIVE = """【学习目标】
- 认识本单元 5 个核心词汇
讲：先带读两遍，再开火车。
【重点词汇】
- tall / short / thin
【句型操练】
- What does he look like?
【作业】
- 抄写单词三遍
"""

SPECIAL = """【符号课】
- 用 & 连起来，<b>不是标签</b>，"双引号"和'单引号'都原样显示。
- 小于号 < 大于号 > 也不能吞掉。
"""


def main():
    print("[1] 5 页 → 5 个 section、页码、标题里的 <script> 被转义")
    output = os.path.join(WORK, "five.html")
    result = build_deck_html({"text": FIVE, "outputPath": output,
                              "title": "词汇课 <script>alert(1)</script> 第一课时"})
    check(result["slides"] == 5, "5 页", result["slides"])
    check(result["size"] == "16:9", "16:9", result["size"])
    raw = open(output, encoding="utf-8").read()
    check(raw.count('<section class="slide') == 5, "5 个 <section class=\"slide\">",
          raw.count('<section class="slide'))
    check(raw.count('class="slide cover"') == 1, "封面带 cover，且只有一个", raw.count('class="slide cover"'))
    check('<div class="pageno">1 / 5</div>' in raw, "首页页码 1 / 5")
    check('<div class="pageno">5 / 5</div>' in raw, "末页页码 5 / 5")
    check("&lt;script&gt;alert(1)&lt;/script&gt;" in raw, "标题里的 <script> 被转义成文本")
    check("<script>alert(1)</script>" not in raw, "原文的 <script> 没有漏进 HTML")
    check(result["bytes"] == os.path.getsize(output), "bytes 就是文件字节数",
          (result["bytes"], os.path.getsize(output)))
    check(result["bytes"] == len(raw.encode("utf-8")), "UTF-8 字节数对得上")
    check(result["warnings"] == [], "这份内容没有告警", result["warnings"])

    print("[2] 讲稿进 aside 且默认隐藏，按 N 才出来")
    check('<aside class="notes">' in raw, "有 notes 的页带 aside")
    check("先带读两遍" in raw, "讲稿正文在 HTML 里")
    check(re.search(r"\.notes\s*\{[^}]*display:\s*none", raw) is not None,
          "CSS 里 .notes 默认 display:none")
    check("body.show-notes .notes" in raw, "body.show-notes 才显示")
    check('classList.toggle("show-notes")' in raw, "脚本里 N 键切换 show-notes")
    plain = render_html([{"title": "只有标题", "bullets": []}])
    check('<aside class="notes">' not in plain, "没有讲稿的页不出 aside")

    print("[3] 打印：@page 297×167mm、每页 page-break-after: always、藏页码导航")
    print_block = raw[raw.index("@media print"):]
    check("@page" in print_block, "打印块里有 @page")
    check("size: 297mm 167mm" in print_block, "@page 尺寸 297mm 167mm（A4 横放且 16:9）")
    check("page-break-after: always" in print_block, "每页 page-break-after: always")
    check("display: flex !important" in print_block, "打印时所有页都显示（不再 display:none）")
    check(".controls" in print_block and ".pageno" in print_block and "display: none !important" in print_block,
          "打印时页码/导航被隐藏")
    check(".slide:last-of-type" in print_block, "最后一页不再断页（不多吐空白纸）")
    check("transform: none !important" in print_block, "打印时取消缩放，按纸张排")

    print("[4] 单文件：没有任何外链，<script> 只有内联一段")
    check("http://" not in raw, "没有 http://")
    check("https://" not in raw, "没有 https://")
    check("//cdn" not in raw, "没有 //cdn")
    check("<link" not in raw, "没有 <link>（不引外部样式）")
    check("src=" not in raw, "没有 src= 外部资源")
    check("@import" not in raw, "没有 @import")
    check(raw.count("<script") == 1, "<script> 只有一段", raw.count("<script"))
    check("</script>" in raw and raw.count("</script>") == 1, "只有一段闭合的 script")
    check(raw.startswith("<!DOCTYPE html>"), "完整文档：有 doctype")
    check('<meta charset="utf-8">' in raw, "声明 UTF-8")

    print("[5] 空 text 走 lesson 的 text-required，不被吞掉")
    expect_error(lambda: build_deck_html({"text": "   ", "outputPath": os.path.join(WORK, "empty.html")}),
                 "text-required", "全空白 → text-required")
    expect_error(lambda: build_deck_html({}), "text-required", "没给 text → text-required")
    check(not os.path.exists(os.path.join(WORK, "empty.html")), "报错时不留半截文件")

    print("[6] 100 条要点不崩，文件不超 400KB")
    hundred = "【要点】\n" + "\n".join("- 第 %d 条要点内容，写长一点看看会不会崩。" % i
                                      for i in range(1, 101))
    big_output = os.path.join(WORK, "big.html")
    big = build_deck_html({"text": hundred, "outputPath": big_output, "title": "一百条要点"})
    big_raw = open(big_output, encoding="utf-8").read()
    check(big["slides"] >= 17, "100 条按 6 条一页切开", big["slides"])
    check(big_raw.count('class="bullet level-0"') >= 100, "100 条一条不少",
          big_raw.count('class="bullet level-0"'))
    check(big["bytes"] < 400 * 1024, "文件 < 400KB", big["bytes"])
    check(big["bytes"] == os.path.getsize(big_output), "bytes 对得上")

    print("[7] & < > 引号与中文标题都正确")
    special = build_deck_html({"text": SPECIAL, "title": "符号 & <课件> 测验",
                               "outputPath": os.path.join(WORK, "special.html")})
    special_raw = open(special["outputPath"], encoding="utf-8").read()
    check("&amp;" in special_raw, "& 转成 &amp;")
    check("&lt;b&gt;" in special_raw, "<b> 转成 &lt;b&gt;")
    check("&gt;" in special_raw, "> 转成 &gt;")
    check("&quot;" in special_raw, '" 转成 &quot;')
    check("&#x27;" in special_raw, "' 转成 &#x27;")
    check("<b>不是标签</b>" not in special_raw, "原文标签没有漏进来")
    check("符号 &amp; &lt;课件&gt; 测验" in special_raw, "中文标题原样保留、符号被转义")
    check("大于号 &gt; 也不能吞掉" in special_raw, "正文里的尖括号按文本显示")

    print("[8] 边界：空 slides、字符串要点、层级收口、重复输出")
    expect_error(lambda: render_html([]), "slides-required", "空 slides → slides-required")
    expect_error(lambda: render_html(None), "slides-required", "None → slides-required")
    loose = render_html(["只有一页"], title="散装")
    check(loose.count('<section class="slide') == 1, "字符串页也能渲染")
    check("只有一页" in loose, "字符串当标题")
    strings = render_html([{"title": "字符串要点", "bullets": ["甲", "乙"]}])
    check(strings.count('class="bullet level-0"') == 2, "要点给字符串也接住")
    deep = render_html([{"title": "深层级", "bullets": [{"text": "缩进", "level": 5}]}])
    check('class="bullet level-2"' in deep, "层级收在 2 级，不缩到没地方写")
    twice = os.path.join(WORK, "twice.html")
    build_deck_html({"text": FIVE, "outputPath": twice, "title": "同一份"})
    expect_error(lambda: build_deck_html({"text": FIVE, "outputPath": twice, "title": "同一份"}),
                 "output-exists", "已存在且没给 overwrite → output-exists")
    again = build_deck_html({"text": FIVE, "outputPath": twice, "title": "同一份", "overwrite": True})
    check(again["bytes"] > 0, "overwrite=True 覆盖成功", again["bytes"])
    auto = build_deck_html({"text": FIVE, "title": "自动存"})
    check(auto["outputPath"].endswith(".html"), "自动路径是 .html", auto["outputPath"])
    check(auto["outputPath"].startswith(os.path.join(WORK, "教师助手")), "落在 文档/教师助手/课件 下",
          auto["outputPath"])
    check(os.path.exists(auto["outputPath"]), "自动路径真的写出来了")

    print("[9] 放映脚本：hash 直达、键盘、缩放")
    check("replaceState" in raw, "翻页用 replaceState 更新 #页码")
    check("hashchange" in raw, "手改地址栏 #7 也能跳到那页")
    check('key === "ArrowRight"' in raw and 'key === "PageDown"' in raw,
          "→ / PageDown 下一页")
    check('key === "Home"' in raw and 'key === "End"' in raw, "Home / End 首末页")
    check('key === " "' in raw, "空格翻页")
    check('setProperty("--scale"' in raw, "resize 重算缩放比例")
    check("scale(var(--scale" in raw, "slide 用 transform: scale() 整体缩放")
    check("window.addEventListener(\"resize\", fit)" in raw, "窗口变化时重新适配")

    print()
    if FAILS:
        print("失败 " + str(len(FAILS)) + " 项：")
        for label in FAILS:
            print(" - " + label)
        sys.exit(1)
    print("全部通过 ✅")


if __name__ == "__main__":
    main()

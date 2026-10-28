#!/usr/bin/env python
"""引擎自检：不依赖 pytest，直接 python test_engine.py。"""

import json
import os
import subprocess
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
ENGINE = os.path.dirname(HERE)
sys.path.insert(0, ENGINE)

from teacher_kit.commands import TABLE, docx_new, inspect, papers, pptx_new, selfcheck  # noqa: E402
from teacher_kit.errors import EngineError  # noqa: E402

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
        return
    check(False, label, "没有抛错")


def main():
    work = tempfile.mkdtemp(prefix="teacher-kit-test-")

    print("[1] 库体检")
    report = selfcheck({})
    check(report["ok"] is True, "selfcheck ok", report)
    check(report.get("python-docx") is not None, "python-docx 可用", report.get("python-docx"))
    check(report.get("python-pptx") is not None, "python-pptx 可用", report.get("python-pptx"))
    check("docx:new" in TABLE and "pptx:new" in TABLE, "命令表里有 docx:new / pptx:new")

    print("[2] 默认造卷：中国老师拿到的是 A4，不是 Letter")
    exam_path = os.path.join(work, "试卷.docx")
    built = docx_new({
        "outputPath": exam_path,
        "blocks": [
            {"type": "paragraph", "align": "center", "runs": [{"text": "七年级英语期中测试卷", "bold": True, "sizePt": 16}]},
            {"type": "paragraph", "text": "听力部分（20 分）"},
            {"type": "paragraph", "text": "Ⅰ. 听录音，选择正确的图片。"},
            {"type": "table", "rows": [["题号", "1", "2", "3"], ["答案", "A", "B", "C"]]},
            {"type": "pageBreak"},
            {"type": "paragraph", "text": "参考答案"},
        ],
        "properties": {"title": "七年级英语期中测试卷", "author": "七年级英语备课组"},
    })
    check(os.path.isfile(built["outputPath"]), "docx 落盘", built)
    check(built["blocks"] == 6, "六个内容块都写了", built["blocks"])
    check(built["page"]["size"] == "A4", "默认纸张 A4", built["page"])

    facts = inspect({"path": exam_path})
    check(facts["kind"] == "docx", "识别为 docx")
    primary = facts["primary"]
    check(abs(primary["pageWidthMm"] - 210) < 1, "A4 宽 210mm", primary["pageWidthMm"])
    check(abs(primary["pageHeightMm"] - 297) < 1, "A4 高 297mm", primary["pageHeightMm"])
    check(primary["orientation"] == "纵向", "纵向", primary["orientation"])
    check(abs(primary["marginLeftMm"] - 31.8) < 1, "默认左边距 31.8mm", primary["marginLeftMm"])
    check(facts["assets"]["tables"] == 1, "数到 1 张表", facts["assets"])
    text_all = " ".join(item["text"] for item in facts["preview"])
    check("七年级英语期中测试卷" in text_all, "预览里有卷名")
    check(facts["metadata"]["title"] == "七年级英语期中测试卷", "文档属性写入", facts["metadata"])
    check(facts["metadata"]["author"] == "七年级英语备课组", "作者写入", facts["metadata"])
    check(primary.get("columns") == 1, "默认单栏", primary.get("columns"))

    print("[3] 8 开横排双栏（真试卷的样子）")
    exam8k = os.path.join(work, "8开试卷.docx")
    built8k = docx_new({
        "outputPath": exam8k,
        "page": {"size": "8K", "orientation": "landscape", "columns": 2},
        "blocks": [{"type": "paragraph", "text": "一、单项选择"}],
    })
    check(built8k["page"]["widthMm"] == 368, "8K 横排宽 368mm", built8k["page"])
    check(built8k["page"]["heightMm"] == 260, "8K 横排高 260mm", built8k["page"])
    facts8k = inspect({"path": exam8k})
    check(abs(facts8k["primary"]["pageWidthMm"] - 368) < 1, "读回宽 368mm", facts8k["primary"]["pageWidthMm"])
    check(abs(facts8k["primary"]["marginLeftMm"] - 27) < 1, "试卷左边距 27mm", facts8k["primary"]["marginLeftMm"])
    check(facts8k["primary"].get("columns") == 2, "读回两栏", facts8k["primary"].get("columns"))
    check(facts8k["primary"]["orientation"] == "横向", "读回横向", facts8k["primary"]["orientation"])

    print("[4] 套模板：保留页眉页脚与页边距")
    template_path = os.path.join(work, "学校模板.docx")
    import docx as docx_module
    from docx.shared import Mm
    template = docx_module.Document()
    section = template.sections[0]
    section.page_width = Mm(210)
    section.page_height = Mm(297)
    section.left_margin = Mm(27)
    section.right_margin = Mm(27)
    section.top_margin = Mm(24)
    section.bottom_margin = Mm(24)
    section.header.paragraphs[0].text = "XX 中学 2025—2026 学年第二学期"
    section.footer.paragraphs[0].text = "第 X 页 共 Y 页　　使用班级：______　　出卷老师：______"
    template.add_paragraph("这里是模板里原有的占位内容，应该被清掉")
    template.save(template_path)

    from_template = docx_new({
        "outputPath": os.path.join(work, "套模板.docx"),
        "baseTemplate": template_path,
        "blocks": [{"type": "paragraph", "text": "套模板后的正文"}],
    })
    check(from_template["fromTemplate"] is True, "标记为从模板生成")
    check(from_template["page"] is None, "没显式给纸张就不动模板版式", from_template["page"])
    facts2 = inspect({"path": from_template["outputPath"]})
    check(abs(facts2["primary"]["marginLeftMm"] - 27) < 1, "左边距 27mm 被保留", facts2["primary"]["marginLeftMm"])
    check("XX 中学" in (facts2["primary"]["headerText"] or ""), "页眉被保留", facts2["primary"]["headerText"])
    check("出卷老师" in (facts2["primary"]["footerText"] or ""), "页脚被保留", facts2["primary"]["footerText"])
    body_text = " ".join(item["text"] for item in facts2["preview"])
    check("占位内容" not in body_text, "模板占位正文被清掉", body_text)
    check("套模板后的正文" in body_text, "新正文写进去了", body_text)

    print("[5] 造课件：默认 16:9")
    deck_path = os.path.join(work, "Unit3.pptx")
    deck_built = pptx_new({
        "outputPath": deck_path,
        "slides": [
            {"title": "Unit 3 How do you get to school?", "subtitle": "七年级英语 · 第一课时", "layout": "title"},
            {"title": "Teaching aims", "bullets": ["ask and answer about transport", "listen for key words"]},
            {"title": "New words", "bullets": [{"text": "subway", "level": 0}, {"text": "by bike", "level": 1}]},
        ],
        "properties": {"title": "Unit 3", "author": "七年级英语备课组"},
    })
    check(deck_built["slideSize"] == "16:9", "默认 16:9", deck_built["slideSize"])
    check(deck_built["slides"] == 3, "三页", deck_built["slides"])
    deck = inspect({"path": deck_path})
    check(deck["kind"] == "pptx", "识别为 pptx")
    check(deck["slideCount"] == 3, "读回 3 页", deck["slideCount"])
    check(abs(deck["slideSizeMm"]["width"] - 338.7) < 1, "16:9 宽 338.7mm", deck["slideSizeMm"])
    check("Unit 3" in " ".join(deck["slides"][0]["texts"]), "第一页标题读到了", deck["slides"][0]["texts"])
    check(any("subway" in text for text in deck["slides"][2]["texts"]), "第三页要点读到了", deck["slides"][2]["texts"])
    check(deck["metadata"]["author"] == "七年级英语备课组", "课件作者写入", deck["metadata"])

    print("[6] 纸张清单（面板下拉框用）")
    listing = papers({})
    names = [item["size"] for item in listing["papers"]]
    check("8K" in names and "16K" in names and "A4" in names, "有 8K/16K/A4", names)
    check(listing["defaultSlideSize"] == "16:9", "默认画布 16:9", listing["defaultSlideSize"])
    check(any(item["size"] == "8K" and item["defaultMarginsMm"][2] == 27.0 for item in listing["papers"]), "8K 默认左右 27mm")

    print("[7] 错误形状")
    expect_error(lambda: inspect({"path": os.path.join(work, "不存在.docx")}), "file-not-found", "文件不存在 → file-not-found")
    bogus = os.path.join(work, "x.txt")
    open(bogus, "w", encoding="utf-8").write("hi")
    expect_error(lambda: inspect({"path": bogus}), "unsupported-format", "txt → unsupported-format")
    expect_error(lambda: docx_new({"outputPath": exam_path, "blocks": [{"type": "paragraph", "text": "x"}]}), "output-exists", "目标已存在 → 不覆盖")
    expect_error(lambda: docx_new({"blocks": []}), "output-required", "没给输出路径 → output-required")
    expect_error(lambda: docx_new({"outputPath": os.path.join(work, "y.docx"), "page": {"size": "B4"}}), "unknown-paper", "不认识的纸 → unknown-paper")
    expect_error(lambda: docx_new({"outputPath": os.path.join(work, "z.docx"), "page": {"size": "A4", "marginsMm": {"left": 120, "right": 120}}}), "margin-too-wide", "页边距过大 → margin-too-wide")
    expect_error(lambda: pptx_new({"outputPath": os.path.join(work, "空.pptx"), "slides": []}), "empty-deck", "没有页 → empty-deck")

    print("[8] 页脚占位符认别名：老师/学校模板写「共 {{total}} 页」也得变成真页码域")
    # 学校画像里存的是「第{{page}}页 共{{total}}页」这种写法。以前只认 {{pages}}，
    # {{total}} 会原样印在页脚上，印前检查判 fail、连 PDF 都导不出来。
    alias_doc = os.path.join(work, "alias-footer.docx")
    docx_new({"outputPath": alias_doc, "footer": "第 {{page}} 页 共 {{total}} 页",
              "blocks": [{"type": "paragraph", "text": "正文"}]})
    from docx import Document as DocxDocument  # 这个测试文件别处不需要 python-docx，就近导入

    alias_document = DocxDocument(alias_doc)
    alias_xml = alias_document.sections[0].footer._element.xml
    check("NUMPAGES" in alias_xml, "{{total}} 变成了真总页数域（NUMPAGES）", alias_xml[:300])
    check("PAGE" in alias_xml, "{{page}} 变成了真页码域（PAGE）", alias_xml[:300])
    alias_text = "\n".join(paragraph.text for paragraph in alias_document.sections[0].footer.paragraphs)
    check("{{" not in alias_text, "页脚里没有剩下的花括号占位符", alias_text)

    print("[9] CLI 协议（宿主就是这么调的）")
    cli = os.path.join(ENGINE, "cli.py")
    good = subprocess.run([sys.executable, cli, "inspect", "--args", json.dumps({"path": exam_path})], capture_output=True, text=True, encoding="utf-8")
    check(good.returncode == 0, "CLI 退出码 0", good.returncode)
    lines = [line for line in good.stdout.strip().splitlines() if line.strip()]
    check(len(lines) == 1, "stdout 只有一行 JSON", good.stdout[:200])
    parsed = json.loads(lines[0])
    check(parsed.get("ok") is True, "JSON ok=true", parsed.get("ok"))
    check(parsed["data"]["kind"] == "docx", "带数据")

    bad = subprocess.run([sys.executable, cli, "inspect", "--args", json.dumps({"path": "./nope.docx"})], capture_output=True, text=True, encoding="utf-8")
    check(bad.returncode == 2, "业务错误退出码 2", bad.returncode)
    check(json.loads(bad.stdout.strip())["error"]["code"] == "file-not-found", "错误 JSON 可解析")
    check(bad.stderr.strip() == "", "业务错误不往 stderr 吐栈", bad.stderr[:120])

    unknown = subprocess.run([sys.executable, cli, "什么鬼"], capture_output=True, text=True, encoding="utf-8")
    check(json.loads(unknown.stdout.strip())["error"]["code"] == "unknown-command", "未知命令有中文报错")

    stdin_case = subprocess.run([sys.executable, cli, "selfcheck", "--args", "-"], input="{}", capture_output=True, text=True, encoding="utf-8")
    check(json.loads(stdin_case.stdout.strip())["ok"] is True, "--args - 从 stdin 读")

    print("")
    if FAILS:
        print(str(len(FAILS)) + " 项失败 ❌")
        for item in FAILS:
            print("  - " + item)
        return 1
    print("全部通过 ✅")
    return 0


if __name__ == "__main__":
    sys.exit(main())

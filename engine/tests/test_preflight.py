#!/usr/bin/env python
"""印前检查器：跑 python test_preflight.py。"""

import os
import sys
import tempfile
import time

import docx
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

HERE = os.path.dirname(os.path.abspath(__file__))
ENGINE = os.path.dirname(HERE)
sys.path.insert(0, ENGINE)

from teacher_kit.commands import TABLE  # noqa: E402
from teacher_kit.errors import EngineError  # noqa: E402
from teacher_kit.placeholders import fill_paragraph  # noqa: E402
from teacher_kit.preflight import preflight  # noqa: E402

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


def build(work, name, page, blocks, footer=True, footer_text="第 {{page}} 页 共 {{pages}} 页 · 使用班级：七(3)班 · 出卷老师：李老师",
          default_footer=True):
    path = os.path.join(work, name)
    TABLE["docx:new"]({"outputPath": path, "page": page, "blocks": blocks, "defaultFooter": default_footer})
    if footer and footer_text is not None:
        document = docx.Document(path)
        paragraph = document.sections[0].footer.paragraphs[0]
        paragraph.text = footer_text
        fill_paragraph(paragraph, {"page": "", "pages": ""})
        document.save(path)
    return path


def set_footer_xml(path, body):
    """把页脚换成手写的域 XML（模拟 Word 里各种页码域的写法）。"""
    document = docx.Document(path)
    paragraph = document.sections[0].footer.paragraphs[0]
    new_p = parse_xml("<w:p " + nsdecls("w") + ">" + body + "</w:p>")
    paragraph._p.addnext(new_p)
    paragraph._p.getparent().remove(paragraph._p)
    document.save(path)
    return path


def clean_blocks():
    return [
        {"type": "paragraph", "align": "center", "runs": [{"text": "七年级英语期中测试卷", "bold": True, "sizePt": 16}]},
        {"type": "paragraph", "align": "center", "text": "（满分 100 分，时间 90 分钟）"},
        {"type": "paragraph", "text": "一、听力（20 分）"},
        {"type": "paragraph", "text": "1. 听录音，选择正确的图片。"},
        {"type": "paragraph", "text": "2. 听录音，补全句子。"},
        {"type": "paragraph", "text": "二、单项选择（30 分）"},
        {"type": "paragraph", "text": "3. — How are you? — ______"},
        {"type": "paragraph", "text": "三、完形填空（20 分）"},
        {"type": "paragraph", "text": "4. 阅读短文，选出最佳答案。"},
        {"type": "paragraph", "text": "四、阅读理解（30 分）"},
        {"type": "paragraph", "text": "5. 判断正误。"},
    ]


def levels(result):
    return [item["level"] for item in result["findings"]]


def codes(result):
    return [item["code"] for item in result["findings"]]


def main():
    work = tempfile.mkdtemp(prefix="teacher-kit-preflight-")

    print("[1] 一份规矩的 8 开两栏卷子：应当直接通过")
    good = build(work, "good.docx", {"size": "8K", "orientation": "landscape", "columns": 2}, clean_blocks())
    result = preflight({"path": good})
    print("     " + result["summary"] + " · " + str(result["score"]) + " 分 · " + "、".join(codes(result)))
    check(result["ok"] is True, "ok = True", levels(result))
    check(result["level"] == "pass", "level = pass", result["findings"])
    check(result["facts"]["paper"] == "8K", "认出 8 开", result["facts"].get("paper"))
    check(result["facts"]["footer"]["hasPageField"] is True, "页脚页码是真域", result["facts"]["footer"])
    check(result["facts"]["score"]["sum"] == 100.0, "分值合计 100", result["facts"]["score"])
    check("footer-no-page-field" not in codes(result), "没有误报页码域")

    print("[2] 纸型/版式/页脚都不对：要拦下来（error），并且给出可改的建议")
    bad = build(
        work,
        "bad.docx",
        {"size": "A4", "orientation": "portrait", "marginsMm": {"top": 10, "bottom": 10, "left": 10, "right": 10}},
        [
            {"type": "paragraph", "align": "center", "text": "{{examTitle}}"},
            {"type": "paragraph", "text": "（满分 100 分，时间 90 分钟）"},
            {"type": "paragraph", "text": "一、听力（20 分）"},
            {"type": "paragraph", "text": "1. 第一题"},
            {"type": "paragraph", "text": "3. 第三题"},
            {"type": "paragraph", "text": "二、单项选择（30 分）"},
        ],
        footer=False,
        default_footer=False,
    )
    result = preflight({"path": bad})
    found = codes(result)
    print("     " + result["summary"] + " · " + str(result["score"]) + " 分 · " + "、".join(found))
    check(result["ok"] is False, "ok = False")
    check(result["level"] == "fail", "level = fail", found)
    check("placeholders-left" in found, "抓到没填的占位符")
    check("numbering-gap-small" in found, "抓到小题跳号（1 → 3）")
    check("footer-missing" in found, "抓到没有页脚")
    check("margin-off" in found, "抓到页边距不对")
    check("score-mismatch" in found, "抓到分值对不上（50 vs 100）")
    check(any(item.get("hint") for item in result["findings"]), "每条问题都带改法")

    print("[3] 8 开横向只有一栏：提醒但不拦")
    single = build(work, "single.docx", {"size": "8K", "orientation": "landscape", "columns": 1}, clean_blocks())
    result = preflight({"path": single})
    check("paper-columns" in codes(result), "提示单栏", codes(result))
    check(result["ok"] is True, "不拦导出", levels(result))

    print("[4] 页脚是手打页码而不是域：必须 error（改内容后页码不会跟着变）")
    typed = build(work, "typed.docx", {"size": "8K", "orientation": "landscape", "columns": 2}, clean_blocks(), footer_text="第 1 页 共 2 页 · 七(3)班")
    result = preflight({"path": typed})
    check("footer-no-page-field" in codes(result), "抓到手打页码", codes(result))
    check(result["level"] == "fail", "level = fail")

    print("[5] 答题框编号不连续：提醒（学生涂卡会对不上）")
    table_path = build(
        work,
        "table.docx",
        {"size": "8K", "orientation": "landscape", "columns": 2},
        [
            {"type": "paragraph", "text": "一、单项选择（100 分）"},
            {"type": "table", "rows": [["题号", "1", "2", "4"], ["答案", "", "", ""]]},
        ],
    )
    result = preflight({"path": table_path})
    check("answer-table-numbering" in codes(result), "抓答题框编号", codes(result))
    check(result["facts"].get("answerTableNumbers") == [1, 2, 4], "读出答题框题号", result["facts"].get("answerTableNumbers"))

    print("[6] 分值合计与满分一致时不报警；不一致时报警")
    check("score-mismatch" not in codes(preflight({"path": good})), "一致的 100 分不报警")
    mismatch = build(
        work,
        "mismatch.docx",
        {"size": "8K", "orientation": "landscape", "columns": 2},
        [
            {"type": "paragraph", "align": "center", "text": "（满分 120 分，时间 120 分钟）"},
            {"type": "paragraph", "text": "一、听力（20 分）"},
            {"type": "paragraph", "text": "二、笔试（30 分）"},
        ],
    )
    result = preflight({"path": mismatch})
    item = next((x for x in result["findings"] if x["code"] == "score-mismatch"), None)
    check(item is not None, "分值不一致报警", codes(result))
    check(item is not None and "50" in item["message"], "报出实际合计 50 分", item)

    print("[7] 传了 values.totalScore 就按它比")
    result = preflight({"path": good, "values": {"totalScore": "120"}})
    check("score-mismatch" in codes(result), "按 values 里的满分比", codes(result))

    print("[8] 错误形状：文件不存在 / 没给路径")
    expect_error(lambda: preflight({}), "path-required", "没给路径 → path-required")
    expect_error(lambda: preflight({"path": os.path.join(work, "nope.docx")}), "file-not-found", "文件不存在 → file-not-found")

    print("[9] 每大题内小题从 1 重新编号：不能误报（国内初中卷的常态）")
    restart = build(
        work,
        "restart.docx",
        {"size": "8K", "orientation": "landscape", "columns": 2},
        [
            {"type": "paragraph", "align": "center", "text": "（满分 100 分，时间 90 分钟）"},
            {"type": "paragraph", "text": "一、听力（20 分）"},
            {"type": "paragraph", "text": "1. 听句子选图。"},
            {"type": "paragraph", "text": "2. 听对话答题。"},
            {"type": "paragraph", "text": "3. 听短文填空。"},
            {"type": "paragraph", "text": "二、单项选择（30 分）"},
            {"type": "paragraph", "text": "1. — How are you? — ______"},
            {"type": "paragraph", "text": "2. This is ______ apple."},
            {"type": "paragraph", "text": "三、书面表达（50 分）"},
            {"type": "paragraph", "text": "1. 以 My Day 为题写一段话。"},
        ],
    )
    result = preflight({"path": restart})
    numbering = [code for code in codes(result) if code.startswith("numbering-")]
    check(not numbering, "分段重编号不误报", numbering)
    check(result["ok"], "分段重编号仍可导出", result["summary"])

    print("[10] 日期和小数不能被当成题号")
    dates = build(
        work,
        "dates.docx",
        {"size": "8K", "orientation": "landscape", "columns": 2},
        [
            {"type": "paragraph", "align": "center", "text": "（满分 100 分，时间 90 分钟）"},
            {"type": "paragraph", "text": "一、听力（100 分）"},
            {"type": "paragraph", "text": "1. 听句子选图。"},
            {"type": "paragraph", "text": "2. 听对话答题。"},
            {"type": "paragraph", "text": "3. 听短文填空。"},
            {"type": "paragraph", "text": "10.15 前交卷。"},
            {"type": "paragraph", "text": "3.14 是圆周率。"},
            {"type": "paragraph", "text": "1.5 倍速播放。"},
            {"type": "paragraph", "text": "9.1 分也及格。"},
        ],
    )
    result = preflight({"path": dates})
    gaps = [code for code in codes(result) if code.startswith("numbering-gap")]
    check(not gaps, "日期/小数不当题号（不报跳号）", gaps)

    print("[11] 页码必须是真域：手打「第 1 页」+ NUMPAGES 域也得拦")
    typed_page = build(
        work, "typed-page.docx",
        {"size": "8K", "orientation": "landscape", "columns": 2},
        clean_blocks(), footer=False,
    )
    set_footer_xml(
        typed_page,
        '<w:r><w:t>第 1 页 共 </w:t></w:r>'
        '<w:fldSimple w:instr=" NUMPAGES "><w:r><w:t>6</w:t></w:r></w:fldSimple>'
        '<w:r><w:t> 页 · 使用班级：七(3)班 · 出卷老师：李老师</w:t></w:r>',
    )
    result = preflight({"path": typed_page})
    check("footer-no-page-field" in codes(result), "手打页码 + NUMPAGES 域 → 报错", codes(result))
    check(not result["ok"], "手打页码拦导出")

    section_pages = build(
        work, "section-pages.docx",
        {"size": "8K", "orientation": "landscape", "columns": 2},
        clean_blocks(), footer=False,
    )
    set_footer_xml(
        section_pages,
        '<w:fldSimple w:instr=" SECTIONPAGES "><w:r><w:t>6</w:t></w:r></w:fldSimple>'
        '<w:r><w:t> 页 · 使用班级：七(3)班 · 出卷老师：李老师</w:t></w:r>',
    )
    check(
        "footer-no-page-field" in codes(preflight({"path": section_pages})),
        "SECTIONPAGES 不算页码域",
        codes(preflight({"path": section_pages})),
    )

    print("[12] 老式 fldChar 域、instrText 被拆成两段：都要认得出")
    old_field = build(
        work, "old-field.docx",
        {"size": "8K", "orientation": "landscape", "columns": 2},
        clean_blocks(), footer=False,
    )
    set_footer_xml(
        old_field,
        '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
        '<w:r><w:instrText> PAGE </w:instrText></w:r>'
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
        '<w:r><w:t>1</w:t></w:r>'
        '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
        '<w:r><w:t> 页 · 使用班级：七(3)班 · 出卷老师：李老师</w:t></w:r>',
    )
    result = preflight({"path": old_field})
    check("footer-no-page-field" not in codes(result), "老式 fldChar 域认得出", codes(result))
    check(result["facts"]["footer"]["hasPageField"], "老式域的 hasPageField = true")

    split_field = build(
        work, "split-field.docx",
        {"size": "8K", "orientation": "landscape", "columns": 2},
        clean_blocks(), footer=False,
    )
    set_footer_xml(
        split_field,
        '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
        '<w:r><w:instrText>PA</w:instrText></w:r>'
        '<w:r><w:instrText>GE</w:instrText></w:r>'
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
        '<w:r><w:t>1</w:t></w:r>'
        '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
        '<w:r><w:t> 页 · 使用班级：七(3)班 · 出卷老师：李老师</w:t></w:r>',
    )
    check(
        "footer-no-page-field" not in codes(preflight({"path": split_field})),
        "instrText 拆两段也认得出",
    )

    print("[13] 表格单元格里的占位符也要报")
    token_table = build(
        work, "token-table.docx",
        {"size": "8K", "orientation": "landscape", "columns": 2},
        clean_blocks() + [{"type": "table", "rows": [["使用班级", "{{unheardOfToken}}"]]}],
    )
    check("placeholders-left" in codes(preflight({"path": token_table})), "表格里的未知占位符报 error")

    print("[14] 打不开的文件给统一错误形状")
    broken = os.path.join(work, "broken.docx")
    with open(broken, "wb") as handle:
        handle.write(b"PK not really a docx")
    expect_error(lambda: preflight({"path": broken}), "invalid-docx", "损坏的 docx → invalid-docx")

    print("[15] 上千段的问题卷：同类问题不刷屏、别卡住")
    noisy = build(
        work, "noisy.docx",
        {"size": "8K", "orientation": "landscape", "columns": 2},
        [{"type": "paragraph", "text": "2. 重复的题目"} for _ in range(300)],
    )
    started = time.monotonic()
    result = preflight({"path": noisy})
    elapsed = time.monotonic() - started
    check(len(result["findings"]) <= 30, "同类问题最多留几条", len(result["findings"]))
    check(elapsed < 5.0, "300 段也在 5 秒内", round(elapsed, 2))

    print("[16] 页眉里的下载站广告要提醒（会跟着卷子一起印）")
    ad_path = build(
        work, "ad.docx",
        {"size": "8K", "orientation": "landscape", "columns": 2},
        clean_blocks(),
    )
    ad_doc = docx.Document(ad_path)
    ad_doc.sections[0].header.paragraphs[0].text = "www.xxjyw.com 免费下载 请勿商用"
    ad_doc.save(ad_path)
    ad_result = preflight({"path": ad_path})
    check("header-ad" in codes(ad_result), "页眉广告报 header-ad", codes(ad_result))

    print("[17] 学校抬头不能被当成广告；Letter 纸型要说明白")
    school_path = build(
        work, "school.docx",
        {"size": "8K", "orientation": "landscape", "columns": 2},
        clean_blocks(),
    )
    school_doc = docx.Document(school_path)
    school_doc.sections[0].header.paragraphs[0].text = "实验中学 2026 学年第一学期 七年级英语 期中测试"
    school_doc.save(school_path)
    check("header-ad" not in codes(preflight({"path": school_path})), "学校抬头不报广告")
    letter = build(
        work, "letter.docx",
        {"size": "Letter", "orientation": "portrait", "columns": 1},
        clean_blocks(),
    )
    letter_result = preflight({"path": letter})
    check("paper-letter" in codes(letter_result), "Letter 纸型给一句说明", codes(letter_result))

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

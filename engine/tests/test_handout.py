#!/usr/bin/env python
"""听写/默写单：跑 python test_handout.py。样本按七年级英语老师真实会粘的样子写。"""

import os
import re
import sys
import tempfile
import zipfile

from docx import Document
from docx.oxml.ns import qn
from lxml import etree

HERE = os.path.dirname(os.path.abspath(__file__))
ENGINE = os.path.dirname(HERE)
sys.path.insert(0, ENGINE)

from teacher_kit.commands import TABLE  # noqa: E402
from teacher_kit.errors import EngineError  # noqa: E402
from teacher_kit.handout import build_dictation, parse_items  # noqa: E402

NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
RE_ITEM = re.compile(r"^\d{1,3}\. ")

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


def section_facts(path):
    """读回每节的栏数和纸型（twips）：分栏和 A4 是这份纸不能错的两件事。"""
    with zipfile.ZipFile(path) as archive:
        tree = etree.fromstring(archive.read("word/document.xml"))
    facts = []
    for sect in tree.iter("{%s}sectPr" % NS_W):
        cols = sect.find("{%s}cols" % NS_W)
        pgsz = sect.find("{%s}pgSz" % NS_W)
        facts.append({
            "cols": int((cols.get("{%s}num" % NS_W) if cols is not None else 1) or 1),
            "width": int(pgsz.get("{%s}w" % NS_W) or 0) if pgsz is not None else 0,
            "height": int(pgsz.get("{%s}h" % NS_W) or 0) if pgsz is not None else 0,
        })
    return facts


def spacing_of(paragraph):
    ppr = paragraph._p.find(qn("w:pPr"))
    if ppr is None:
        return {}
    element = ppr.find(qn("w:spacing"))
    if element is None:
        return {}
    return {key: element.get(qn("w:" + key)) for key in ("line", "lineRule", "before", "after")}


def texts_of(document):
    return [paragraph.text for paragraph in document.paragraphs]


def main():
    work = tempfile.mkdtemp()
    # 别把测试产物写进老师真实的「文档/教师助手」
    os.environ["TEACHER_KIT_HOME"] = os.path.join(work, "教师助手")

    words = "\n".join("word" + str(i) + "\t释义" + str(i) for i in range(1, 11))

    print("[1] 10 条 tab 分隔：序号重排、prompt/answer 跟着 mode 翻面")
    parsed = parse_items(words)
    items = parsed["items"]
    check(len(items) == 10, "10 条全部拆出", len(items))
    check([item["index"] for item in items] == list(range(1, 11)), "序号重排成 1..10",
          [item["index"] for item in items])
    check(items[0]["prompt"] == "word1" and items[0]["answer"] == "释义1",
          "en2cn：prompt 取英文、answer 取中文", items[0])
    check(items[9]["prompt"] == "word10" and items[9]["answer"] == "释义10", "第 10 条也对得上", items[9])
    check(not parsed["warnings"], "干净词表不报警告", parsed["warnings"])
    sworn = parse_items(words, "cn2en")["items"]
    check(sworn[0]["prompt"] == "释义1" and sworn[0]["answer"] == "word1",
          "cn2en：prompt 取中文、answer 取英文", sworn[0])
    check([item["given"] for item in items] == [item["prompt"] for item in items], "given 是 prompt 的别名")
    check([item["blank"] for item in items] == [item["answer"] for item in items], "blank 是 answer 的别名")

    print("[2] 无序号 + 各种分隔符：拆得开、拆得对、拆不开的留警告")
    dirty = "\n".join([
        "apple  苹果",          # 两个空格
        "banana，香蕉",          # 全角逗号
        "cherry = 樱桃",         # 等号
        "date — 枣",            # 破折号
        "egg 蛋",               # 一个空格（中英交界）
        "ice-cream 冰淇淋",      # 词内连字符不能切断
        "look after 照顾",       # 英文短语，空格不能当分隔符
        "fig（无花果）",          # 中文括号
        "3. grape\t葡萄",        # 行首带 3. 要剥掉
        "只写中文的条目",          # 纯中文
        "grammar",              # 纯英文
    ])
    parsed = parse_items(dirty)
    items = parsed["items"]
    check(len(items) == 11, "11 条一条不丢", len(items))
    check(items[0]["prompt"] == "apple" and items[0]["answer"] == "苹果", "两个空格能拆", items[0])
    check(items[1]["prompt"] == "banana" and items[1]["answer"] == "香蕉", "全角逗号能拆", items[1])
    check(items[2]["prompt"] == "cherry" and items[2]["answer"] == "樱桃", "等号能拆", items[2])
    check(items[3]["prompt"] == "date" and items[3]["answer"] == "枣", "破折号能拆", items[3])
    check(items[4]["prompt"] == "egg" and items[4]["answer"] == "蛋", "一个空格也能拆", items[4])
    check(items[5]["prompt"] == "ice-cream", "ice-cream 没被连字符切断", items[5])
    check(items[6]["prompt"] == "look after" and items[6]["answer"] == "照顾", "英文短语整体当前缀", items[6])
    check(items[7]["prompt"] == "fig" and items[7]["answer"] == "无花果", "中文括号能拆", items[7])
    check(items[8]["prompt"] == "grape", "行首 3. 被剥掉后 prompt 是 grape", items[8])
    check(items[9]["prompt"] == "只写中文的条目", "纯中文条目仍出一条（prompt 扶正）", items[9])
    codes = [warning["code"] for warning in parsed["warnings"]]
    check("no-chinese" in codes, "纯英文记 no-chinese", parsed["warnings"])
    check("no-english" in codes, "纯中文记 no-english", parsed["warnings"])
    only_english = [w for w in parsed["warnings"] if w["code"] == "no-chinese"]
    check(only_english and only_english[0]["message"] == "第 11 条只有英文，没有中文释义",
          "警告文案点明第几条", only_english)

    print("[3] mix：偶数条给英文、奇数条给中文，防同桌互抄")
    mixed = parse_items("one 一\ntwo 二\nthree 三\nfour 四", "mix")["items"]
    check([item["prompt"] for item in mixed] == ["one", "二", "three", "四"],
          "mix 奇偶交替", [item["prompt"] for item in mixed])
    check([item["answer"] for item in mixed] == ["一", "two", "三", "four"],
          "mix 的答案随之翻面", [item["answer"] for item in mixed])
    expect_error(lambda: parse_items("a 一", "en-cn"), "bad-mode", "方式名写错报 bad-mode")

    print("[4] columns=1/2/3：正文节栏数对得上，纸永远是 A4 竖版")
    for columns in (1, 2, 3):
        output = os.path.join(work, "columns" + str(columns) + ".docx")
        result = build_dictation({"text": words, "outputPath": output, "columns": columns})
        check(result["columns"] == columns, "columns=" + str(columns) + " 回传一致", result["columns"])
        facts = section_facts(output)
        check(facts and facts[-1]["cols"] == columns,
              "columns=" + str(columns) + " 正文章节是 " + str(columns) + " 栏", facts)
        check(all(abs(fact["width"] - 11906) <= 20 and abs(fact["height"] - 16838) <= 20 for fact in facts),
              "columns=" + str(columns) + " 纸张是 A4 11906×16838 twips", facts)
        if columns > 1:
            check(facts[0]["cols"] == 1, "columns=" + str(columns) + " 时标题节仍是单栏（标题不挤进半栏）",
                  facts[0])
        else:
            check(len(facts) == 1, "单栏时不切多余的节", len(facts))

    print("[5] 正文段落：1. + 下划线 + 2 倍行距 + 段后 0")
    output = os.path.join(work, "spacing.docx")
    build_dictation({"text": words, "outputPath": output, "blankLength": 20})
    document = Document(output)
    item_paragraphs = [paragraph for paragraph in document.paragraphs if RE_ITEM.match(paragraph.text)]
    check(len(item_paragraphs) == 10, "10 条各占一段", len(item_paragraphs))
    check(item_paragraphs[0].text.startswith("1. word1"), "第一条以 1. 开头", item_paragraphs[0].text)
    check("_" * 20 in item_paragraphs[0].text, "下划线长度 = blankLength", item_paragraphs[0].text)
    check(item_paragraphs[0].text.count("_") == 20, "没有多余的横线", item_paragraphs[0].text.count("_"))
    facts = spacing_of(item_paragraphs[0])
    check(facts.get("line") == "480", "行距 2.0 落到 w:line=480", facts)
    check(facts.get("lineRule") == "auto", "行距规则是倍数（auto）", facts)
    check(facts.get("after") == "0", "段后 0", facts)
    defaulted = os.path.join(work, "default-blank.docx")
    build_dictation({"text": "apple 苹果", "outputPath": defaulted})
    first = [p for p in Document(defaulted).paragraphs if RE_ITEM.match(p.text)][0]
    check(first.text.count("_") == 14, "不传 blankLength 时默认 14 个下划线", first.text.count("_"))

    print("[6] answer=True 有参考答案且覆盖全部条目；False 时一条不露")
    with_answer = os.path.join(work, "with-answer.docx")
    build_dictation({"text": words, "outputPath": with_answer, "answer": True})
    document = Document(with_answer)
    texts = texts_of(document)
    check("参考答案" in texts, "有「参考答案」页", texts[:4])
    expected = [str(i) + ". word" + str(i) + " — 释义" + str(i) for i in range(1, 11)]
    check(all(line in texts for line in expected), "10 条答案一条不落",
          [line for line in expected if line not in texts])
    check(sum(1 for text in texts if " — " in text) == 10, "答案行数正好等于条目数",
          sum(1 for text in texts if " — " in text))
    heading = [p for p in document.paragraphs if p.text == "参考答案"]
    check(bool(heading) and heading[0].paragraph_format.page_break_before is True,
          "答案页另起一页", heading[0].paragraph_format.page_break_before if heading else None)
    facts = section_facts(with_answer)
    check([fact["cols"] for fact in facts] == [1, 2, 1, 2],
          "答案页结构：标题单栏 / 条目双栏 / 答案标题单栏 / 答案双栏", facts)
    swapped = os.path.join(work, "with-answer-cn2en.docx")
    build_dictation({"text": words, "outputPath": swapped, "answer": True, "mode": "cn2en"})
    check("1. word1 — 释义1" in texts_of(Document(swapped)),
          "cn2en 的答案仍是「英文 — 中文」顺序")
    plain = os.path.join(work, "no-answer.docx")
    build_dictation({"text": words, "outputPath": plain})
    plain_texts = texts_of(Document(plain))
    check("参考答案" not in plain_texts, "answer=False 时没有答案页", plain_texts[-3:])
    check(not [text for text in plain_texts if " — " in text], "answer=False 时没有答案行")

    print("[7] 上限、去重、count、空输入与坏参数")
    many = "\n".join("w" + str(i) + " 词" + str(i) for i in range(1, 131))
    parsed = parse_items(many)
    check(len(parsed["items"]) == 120, "超过 120 条截到 120", len(parsed["items"]))
    check(any(warning["code"] == "too-many" for warning in parsed["warnings"]), "截断给 too-many 警告",
          parsed["warnings"])
    duplicated = parse_items("apple 苹果\nbanana 香蕉\napple 苹果\napple 苹果")
    check(len(duplicated["items"]) == 2, "重复按 prompt 去重、保序", duplicated["items"])
    check(any(warning["code"] == "duplicate" and warning["count"] == 2 for warning in duplicated["warnings"]),
          "重复条数记进警告", duplicated["warnings"])
    limited = os.path.join(work, "count.docx")
    result = build_dictation({"text": words, "outputPath": limited, "count": 5})
    check(result["count"] == 5, "count 只取前 N", result["count"])
    check(len([p for p in Document(limited).paragraphs if RE_ITEM.match(p.text)]) == 5,
          "纸上也只印了 5 条")
    too_many_wanted = build_dictation({"text": "a 一\nb 二", "outputPath": os.path.join(work, "few.docx"),
                                       "count": 9})
    check(any(warning["code"] == "not-enough-items" for warning in too_many_wanted["warnings"]),
          "count 比条目多时给 not-enough-items 而不是报错", too_many_wanted["warnings"])
    expect_error(lambda: build_dictation({"text": "", "outputPath": os.path.join(work, "x.docx")}),
                 "text-required", "空 text 报 text-required")
    expect_error(lambda: build_dictation({"text": "   \n\t\n", "outputPath": os.path.join(work, "x.docx")}),
                 "text-required", "只有空白也报 text-required")
    expect_error(lambda: build_dictation({"text": ",,,", "outputPath": os.path.join(work, "x.docx")}),
                 "no-items", "一条都认不出时报 no-items")
    expect_error(lambda: build_dictation({"text": "a 一", "columns": 4}), "bad-columns", "4 栏拒绝（最多 3）")
    expect_error(lambda: build_dictation({"text": "a 一", "blankLength": 0}), "bad-blank-length",
                 "下划线 0 个拒绝")
    expect_error(lambda: build_dictation({"text": "a 一", "count": 0}), "bad-count", "count=0 拒绝")
    expect_error(lambda: build_dictation({"text": "a 一", "lineSpacing": 0}), "bad-line-spacing",
                 "行距 0 拒绝")
    existing = os.path.join(work, "exists.docx")
    build_dictation({"text": "a 一", "outputPath": existing})
    expect_error(lambda: build_dictation({"text": "a 一", "outputPath": existing}), "output-exists",
                 "同名文件不覆盖（要覆盖得显式传 overwrite）")

    print("[8] 套学校模板：页眉还在，模板正文不残留")
    template = os.path.join(work, "school-template.docx")
    TABLE["docx:new"]({"outputPath": template, "blocks": [{"type": "paragraph", "text": "上周的卷子"}]})
    document = Document(template)
    header = document.sections[0].header
    header.is_linked_to_previous = False
    paragraph = header.paragraphs[0] if header.paragraphs else header.add_paragraph()
    paragraph.text = "七年级英语 听写专用纸"
    document.save(template)
    output = os.path.join(work, "from-template.docx")
    result = build_dictation({"text": "apple 苹果\npear 梨", "outputPath": output,
                              "baseTemplate": template, "answer": True})
    check(result["count"] == 2, "模板上照样出 2 条", result["count"])
    document = Document(output)
    header_text = "".join(paragraph.text for paragraph in document.sections[0].header.paragraphs)
    check("听写专用纸" in header_text, "页眉还在", header_text)
    output_texts = texts_of(document)
    check("上周的卷子" not in output_texts, "模板里的示例正文被清掉", output_texts[:3])
    check(any(text.startswith("1. apple") for text in output_texts), "条目正常落在模板上")
    check("参考答案" in output_texts, "套模板时答案页也在")

    print("[9] 默认标题/说明行/超长条目/默认输出路径")
    sample = os.path.join(work, "plain.docx")
    build_dictation({"text": "apple 苹果", "outputPath": sample})
    document = Document(sample)
    check(document.paragraphs[0].text == "英语听写", "缺省标题是「英语听写」", document.paragraphs[0].text)
    info = document.paragraphs[1].text
    check(all(word in info for word in ("班级：", "姓名：", "得分：")), "说明行含班级/姓名/得分", info)
    check(document.paragraphs[0].paragraph_format.alignment is not None, "标题是居中的")
    long_answer = "很长的释义" * 30
    long_case = os.path.join(work, "long.docx")
    result = build_dictation({"text": "supercalifragilisticexpialidocious " + long_answer,
                              "outputPath": long_case})
    check(result["count"] == 1, "超长条目仍然只算一条", result["count"])
    long_texts = texts_of(Document(long_case))
    check(any(text.startswith("1. supercalifragilisticexpialidocious") for text in long_texts),
          "超长英文前缀没被截断", long_texts[2:4])
    titled = build_dictation({"text": "a 一", "title": "Unit 3 单词默写"})
    check(os.path.exists(titled["outputPath"]), "不传 outputPath 时落到默认目录", titled["outputPath"])
    check("听写" in titled["outputPath"], "默认目录按任务名「听写」分", titled["outputPath"])
    check(titled["outputPath"].endswith(".docx"), "默认扩展名 docx", titled["outputPath"])
    check(titled["bytes"] > 0 and os.path.getsize(titled["outputPath"]) == titled["bytes"],
          "回传的 bytes 和文件对得上", titled["bytes"])
    custom = os.path.join(work, "16k.docx")
    sixteen = build_dictation({"text": "a 一", "outputPath": custom, "page": {"size": "16K"}})
    check(sixteen["page"]["size"] == "16K", "page 覆盖纸型生效", sixteen["page"])
    facts = section_facts(custom)
    check(abs(facts[0]["width"] - 10432) <= 20 and abs(facts[0]["height"] - 14740) <= 20,
          "16K 纸型真的落到文件里", facts)

    print("[10] 词表首行是标题（Unit 2 词汇听写（20 个））时不当第 1 条（persona）")
    titled_list = parse_items("""Unit 2 词汇听写（20 个）
tall 高的
short 矮的
thin 瘦的
""")
    check([item["prompt"] for item in titled_list["items"]] == ["tall", "short", "thin"],
          "标题行没被当词条", [item["prompt"] for item in titled_list["items"]])
    check("skipped-title" in [warning["code"] for warning in titled_list["warnings"]],
          "会告诉老师那行是标题", [warning["code"] for warning in titled_list["warnings"]])
    real_entry = parse_items("""dictation 听写
tall 高的
""")
    check([item["prompt"] for item in real_entry["items"]] == ["dictation", "tall"],
          "「dictation 听写」是真词条，不能跳过", [item["prompt"] for item in real_entry["items"]])

    print("[11] 老师直接选一个词表文件（Word 表格 / Excel / CSV），不用复制粘贴")
    import csv as csv_module
    from openpyxl import Workbook

    from teacher_kit.handout import read_word_list

    folder = tempfile.mkdtemp(prefix="tk-wordlist-")

    words_doc = Document()
    words_doc.add_paragraph("Unit 4 单词表")
    words_table = words_doc.add_table(rows=4, cols=2)
    for row, (en, cn) in zip(words_table.rows, [("apple", "苹果"), ("banana", "香蕉"),
                                                ("orange", "橘子"), ("pear", "梨")]):
        row.cells[0].text = en
        row.cells[1].text = cn
    words_doc.add_paragraph("grape 葡萄")
    words_path = os.path.join(folder, "词表.docx")
    words_doc.save(words_path)

    word_text = read_word_list(words_path)
    check(word_text.splitlines() == ["apple 苹果", "banana 香蕉", "orange 橘子", "pear 梨", "grape 葡萄"],
          "Word 里一张两列表格 + 散装一行都读得出来", word_text)
    check("单词表" not in word_text, "「Unit 4 单词表」这种标题行不当词条", word_text)

    csv_path = os.path.join(folder, "词表.csv")
    with open(csv_path, "w", encoding="utf-8-sig", newline="") as handle:
        writer = csv_module.writer(handle)
        writer.writerow(["英文", "中文"])
        writer.writerow(["dog", "狗"])
        writer.writerow(["cat", "猫"])
    csv_text = read_word_list(csv_path)
    check(csv_text.strip() == "dog 狗\ncat 猫", "CSV 的表头行跳过、剩下两行留着", csv_text)

    book = Workbook()
    sheet = book.active
    sheet.append(["序号", "英文", "中文"])
    sheet.append([1, "tall", "高的"])
    sheet.append([2, "short", "矮的"])
    xlsx_path = os.path.join(folder, "词表.xlsx")
    book.save(xlsx_path)
    xlsx_text = read_word_list(xlsx_path)
    check(xlsx_text.splitlines() == ["tall 高的", "short 矮的"],
          "Excel 的表头行与序号列都跳过（序号不该印成题面「1. 1 tall」）", xlsx_text)

    made = build_dictation({"source": words_path, "title": "词汇听写"})
    check(made["count"] == 5, "build_dictation 只给文件也能出听写单", made["count"])
    check(os.path.exists(made["outputPath"]), "文件落在磁盘上", made["outputPath"])

    try:
        read_word_list(os.path.join(folder, "没有这个文件.docx"))
        check(False, "文件不存在要报错", "没报")
    except EngineError as error:
        check(error.code == "file-not-found", "文件不存在报 file-not-found", error.to_dict())
        check(bool(error.hint), "错误带怎么做的提示", error.to_dict())

    pdf_path = os.path.join(folder, "词表.pdf")
    with open(pdf_path, "wb") as handle:
        handle.write(b"%PDF-1.4")
    try:
        read_word_list(pdf_path)
        check(False, "PDF 词表要说明读不了", "没报")
    except EngineError as error:
        check(error.code == "unsupported-source", "PDF 报 unsupported-source", error.to_dict())
        check("docx" in (error.hint or ""), "提示里给了能用的格式", error.to_dict())

    print("[12] 带 Unit 编号的标题行（Unit 3 单词听写）也不当第 1 条")
    numbered = parse_items("Unit 3 单词听写\n1. tall 高的\n2. short 矮的\n3. thin 瘦的")
    check(len(numbered["items"]) == 3, "只收 3 条词，标题行不算",
          [item["prompt"] for item in numbered["items"]])
    check(numbered["items"][0]["prompt"] == "tall", "第 1 条是 tall", numbered["items"][0])
    check(any(item["code"] == "skipped-title" for item in numbered["warnings"]),
          "跳过标题要告诉老师", numbered["warnings"])
    # 「practice 练习」「test 测试」这种右边就两个字的，是正经词条
    short_gloss = parse_items("practice 练习\ntest 测试")
    check([item["prompt"] for item in short_gloss["items"]] == ["practice", "test"],
          "右边两个字的词条不误跳", [item["prompt"] for item in short_gloss["items"]])

    print("[13] 整张 Word 词表（首段标题 + 两列表格 + 末段落款）只出 14 条（评测 P0-2）")
    WORDS_14 = [("tall", "高的"), ("short", "矮的；短的"), ("thin", "瘦的"), ("heavy", "重的"),
                ("curly", "卷曲的"), ("straight", "直的"), ("medium", "中等的"),
                ("height", "身高；高度"), ("build", "体格；体型"), ("beard", "胡须"),
                ("glasses", "眼镜"), ("blonde", "金黄色的"), ("describe", "描述"),
                ("look", "看起来")]
    real_doc = Document()
    real_doc.add_paragraph("Unit 3 What does he look like? 单词表")
    real_table = real_doc.add_table(rows=len(WORDS_14) + 1, cols=2)
    for row, (en, cn) in zip(real_table.rows, [("英文", "中文")] + WORDS_14):
        row.cells[0].text = en
        row.cells[1].text = cn
    real_doc.add_paragraph("（Unit 3 单词表 · 七年级英语 · 王老师）")
    real_path = os.path.join(work, "Unit3单词表.docx")
    real_doc.save(real_path)

    real_text = read_word_list(real_path)
    real_lines = real_text.splitlines()
    check(len(real_lines) == 14, "14 个词一条不多一条不少（标题/落款没混进来）", real_lines)
    check(real_lines[0] == "tall 高的", "第 1 行就是第一个词 tall", real_lines[:2])
    check(real_lines[-1] == "look 看起来", "最后一行就是最后一个词 look", real_lines[-2:])
    check("单词表" not in real_text and "王老师" not in real_text,
          "标题与落款都不在词条里", real_text.splitlines()[-2:])

    full = build_dictation({"source": real_path, "outputPath": os.path.join(work, "real-all.docx")})
    check(full["count"] == 14, "整张词表出单正好 14 条", full["count"])
    real_items = [p.text for p in Document(full["outputPath"]).paragraphs if RE_ITEM.match(p.text)]
    check(len(real_items) == 14, "纸上也就 14 条（没有第 15、16 条）", real_items[-3:])
    check(real_items[0].startswith("1. tall"), "第 1 条是 tall", real_items[0])
    check(real_items[-1].startswith("14. look"), "第 14 条是 look", real_items[-1])
    check(not [text for text in real_items if "单词表" in text or "王老师" in text],
          "标题/落款没被印成单词", [t for t in real_items if "单词表" in t or "王老师" in t])
    check(not full["warnings"], "干净词表不再报 no-chinese", full["warnings"])
    limited_real = build_dictation({"source": real_path, "count": 10,
                                    "outputPath": os.path.join(work, "real-10.docx")})
    check(limited_real["count"] == 10, "count:10 时单子 10 条", limited_real["count"])
    check(len([p for p in Document(limited_real["outputPath"]).paragraphs if RE_ITEM.match(p.text)]) == 10,
          "count:10 时纸上也 10 条")

    trailing = parse_items("tall 高的\nshort 矮的\n（Unit 3 单词表 · 七年级英语 · 王老师）")
    check([item["prompt"] for item in trailing["items"]] == ["tall", "short"],
          "落款粘成文本时也不当词条", [item["prompt"] for item in trailing["items"]])
    check("skipped-title" in [warning["code"] for warning in trailing["warnings"]],
          "跳过的落款会告诉老师", trailing["warnings"])
    last_entry = parse_items("tall 高的\ntest 测试")
    check([item["prompt"] for item in last_entry["items"]] == ["tall", "test"],
          "末尾的「test 测试」是真词条，不因位置被跳过",
          [item["prompt"] for item in last_entry["items"]])
    named = parse_items("tall 高的\nname 姓名")
    check([item["prompt"] for item in named["items"]] == ["tall", "name"],
          "「name 姓名」带抬头词也仍是词条", [item["prompt"] for item in named["items"]])

    print()
    if FAILS:
        print("失败 " + str(len(FAILS)) + " 项：")
        for label in FAILS:
            print(" - " + label)
        sys.exit(1)
    print("全部通过 ✅")


if __name__ == "__main__":
    main()

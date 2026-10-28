#!/usr/bin/env python
"""组卷：跑 python test_exam.py。样本全按老师真实粘贴的样子写（一行多选项、续行、缺题号）。"""

import os
import sys
import tempfile

import docx
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

HERE = os.path.dirname(os.path.abspath(__file__))
ENGINE = os.path.dirname(HERE)
sys.path.insert(0, ENGINE)

from teacher_kit.commands import TABLE  # noqa: E402
from teacher_kit.errors import EngineError  # noqa: E402
from teacher_kit.exam import build_exam, exam_parse, parse_paper  # noqa: E402

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


ENGLISH = """一、听力（20 分）
1. What does the boy want to buy?（5 分）
A. A book  B. A pen  C. A bag  D. A cup
2. Where are they talking?（5 分）
A. In the library  B. In the shop
C. In the classroom  D. At home
3. 听录音，补全句子。（5 分）
4. 听录音，判断正误。（5 分）
二、单项选择（30 分）
5. — How are you? — ______
A. Fine, thank you  B. I'm ten  C. Yes, I am  D. That's all
6. This is ______ apple.
A. a  B. an  C. the  D. /
7. My sister likes ______ books.
8. He ______ to school by bus every day.
9. ______ is your English teacher?
10. I have two ______.
三、阅读理解（50 分）
11. 阅读短文，选择最佳答案。
【答案】B
12. 阅读短文，回答下面的问题。
答案：He is a doctor.
13. 阅读短文，判断正误。
14. 阅读短文，选择最佳答案。
15. 阅读短文，补全句子。
"""

MATH = """三、解答题（40 分）
（1）计算：2 + 3 × 4
（2）解方程：x + 5 = 12
（3）已知 a = 3，求 2a + 1
"""


def main():
    work = tempfile.mkdtemp(prefix="exam-")
    os.environ["TEACHER_KIT_HOME"] = os.path.join(work, "教师助手")
    print("[1] 英语卷：大题、小题、一行多选项、答案、均摊分值")
    parsed = parse_paper(ENGLISH)
    check(len(parsed["sections"]) == 3, "切成 3 个大题", [s["title"] for s in parsed["sections"]])
    check(parsed["sections"][0]["title"].startswith("听力"), "大题标题去了编号", parsed["sections"][0]["title"])
    check(not parsed["warnings"], "题号连续不报 renumbered", parsed["warnings"])
    first = parsed["sections"][0]["questions"][0]
    check(len(first["options"]) == 4, "一行四个选项全切出来", [o["label"] for o in first["options"]])
    check(first["options"][1]["text"] == "A pen", "选项文字去空白", first["options"][1])
    second = parsed["sections"][0]["questions"][1]
    check(len(second["options"]) == 4 and second["options"][2]["text"] == "In the classroom",
          "跨行选项各自归位", second["options"])
    listening = parsed["sections"][0]
    check([q["score"] for q in listening["questions"]] == [5, 5, 5, 5], "节内分值对上",
          [q["score"] for q in listening["questions"]])
    choice = parsed["sections"][1]
    check(len(choice["questions"]) == 6, "单选 6 题", len(choice["questions"]))
    check(abs(sum(q["score"] for q in choice["questions"]) - 30) < 1e-6, "单选均摊到 30 分",
          sum(q["score"] for q in choice["questions"]))
    check(parsed["sections"][2]["questions"][0]["answer"] == "B", "【答案】解析", parsed["sections"][2]["questions"][0])
    check(parsed["sections"][2]["questions"][1]["answer"] == "He is a doctor.", "答案：解析",
          parsed["sections"][2]["questions"][1])
    check(abs(parsed["totalScore"] - 100) < 1e-6, "总分 100", parsed["totalScore"])

    print("[2] 数学：（1）(2) 编号 + 跳过一、二只给三、")
    parsed = parse_paper(MATH)
    check(len(parsed["sections"]) == 1, "一个大题", len(parsed["sections"]))
    answers = parsed["sections"][0]
    check(len(answers["questions"]) == 3, "括号小题被识别", [q["number"] for q in answers["questions"]])
    check(answers["questions"][0]["stem"].startswith("计算"), "题干切干净", answers["questions"][0]["stem"])

    print("[3] 题号不连续 → 重排并告诉老师")
    parsed = parse_paper("一、选择题（6 分）\n1. 甲（2 分）\n2. 乙（2 分）\n4. 丙（2 分）\n")
    check([q["number"] for q in parsed["sections"][0]["questions"]] == [1, 2, 3], "重排成 1/2/3",
          [q["number"] for q in parsed["sections"][0]["questions"]])
    check(any(w["code"] == "renumbered" for w in parsed["warnings"]), "给了 renumbered 提示", parsed["warnings"])

    print("[4] 没分值也没 totalScore → 用 scorePerQuestion")
    parsed = parse_paper("一、填空（10 分）\n1. 甲\n2. 乙\n")
    check([q["score"] for q in parsed["sections"][0]["questions"]] == [5, 5], "按大题均摊",
          [q["score"] for q in parsed["sections"][0]["questions"]])

    print("[5] 出整卷：8 开两栏 + A4 答题卡 + 答案页")
    output = os.path.join(work, "english.docx")
    result = build_exam({
        "text": ENGLISH,
        "outputPath": output,
        "title": "七年级英语期中测试卷",
        "page": {"size": "8K", "orientation": "landscape", "columns": 2},
        "values": {"schoolYear": "2025—2026", "minutes": "90"},
        "options": {"minutes": 90, "answerCard": "card", "answerSheet": "end"},
    })
    check(os.path.exists(output), "文件落盘", output)
    check(result["questionCount"] == 15, "15 题", result["questionCount"])
    check(len(result["sections"]) == 3, "三节都在", result["sections"])
    check(result["totalScore"] == 100, "总分 100", result["totalScore"])
    document = docx.Document(output)
    # 两栏卷子的卷头写进「第一页页眉」：正文中途改栏数的分节会被 LibreOffice 推到下一页
    # （长题干实测白印一张 8 开），所以整卷只有正文/答题卡/答案三节。
    check(len(document.sections) == 3, "三节：正文 + 答题卡 + 答案（卷头在第一页页眉）", len(document.sections))
    first_header = "\n".join(paragraph.text for paragraph in document.sections[0].first_page_header.paragraphs)
    check("七年级英语期中测试卷" in first_header, "卷名在第一页页眉里（跨整页宽）", first_header)
    check(not any(paragraph.text.strip() == "七年级英语期中测试卷" for paragraph in document.paragraphs),
          "卷名不再留在正文里（留在正文里就得为它切一节，两栏会错位）")
    body = document.sections[0]
    body_cols = body._sectPr.find(qn("w:cols"))
    check(body_cols is not None and body_cols.get(qn("w:num")) == "2", "正文就是两栏（不再靠分节切栏）",
          body_cols.attrib if body_cols is not None else None)
    check(abs(body.page_width.mm - 368) < 2 and abs(body.page_height.mm - 260) < 2, "正文是 8 开横向",
          (body.page_width.mm, body.page_height.mm))
    card = document.sections[1]
    check(abs(card.page_width.mm - 210) < 2 and abs(card.page_height.mm - 297) < 2, "答题卡是 A4 纵向",
          (card.page_width.mm, card.page_height.mm))
    check(len(document.tables) >= 1, "答题卡有表格", len(document.tables))
    card_cols = card._sectPr.find(qn("w:cols"))
    check(card_cols is None or card_cols.get(qn("w:num")) in (None, "1"), "答题卡是单栏（不继承正文两栏）",
          card_cols.attrib if card_cols is not None else None)
    first_table = document.tables[0]
    check(first_table.cell(0, 0).text == "题号" and first_table.cell(0, 1).text == "1", "表头 题号/1",
          first_table.cell(0, 0).text)
    stem = next((p for p in document.paragraphs if p.text.startswith("1. ")), None)
    check(stem is not None and stem.paragraph_format.keep_with_next, "题干和选项拴在一起（不会被栏切开）",
          stem.paragraph_format.keep_with_next if stem is not None else None)
    heading = next((p for p in document.paragraphs if p.text.startswith("一、")), None)
    check(heading is not None and heading.paragraph_format.keep_with_next, "大题标题跟着第一题走",
          heading.paragraph_format.keep_with_next if heading is not None else None)
    codes = [item.get("code") for item in (result["preflight"].get("findings") or [])]
    check("paper-columns" not in codes, "两栏卷子不会被误报「只有 1 栏」", codes)
    text = "\n".join(paragraph.text for paragraph in document.paragraphs)
    check("参考答案" in text, "答案页有标题")
    check("He is a doctor." in text, "答案内容在")
    check("满分 100 分" in first_header, "说明行分值对（跟卷名一起在页眉）", first_header[:80])
    check(isinstance(result["preflight"], dict) and "score" in result["preflight"], "带回印前检查",
          list(result["preflight"].keys()) if isinstance(result["preflight"], dict) else result["preflight"])

    print("[6] 不要答题卡、不要答案页")
    plain = os.path.join(work, "plain.docx")
    result = build_exam({"text": ENGLISH, "outputPath": plain,
                         "options": {"answerCard": "none", "answerSheet": "none"}})
    document = docx.Document(plain)
    check(len(document.sections) == 1, "只有一节（卷头进页眉，不再为标题多切一节）", len(document.sections))
    plain_header = "\n".join(paragraph.text for paragraph in document.sections[0].first_page_header.paragraphs)
    check("班级" in plain_header and "姓名" in plain_header, "班级/姓名/学号那行也跟着进了页眉",
          plain_header[:80])
    check(not document.tables, "没有答题卡表格", len(document.tables))

    print("[7] 面板预览 exam:parse")
    preview = TABLE["exam:parse"]({"text": ENGLISH})
    check(preview["questionCount"] == 15, "预览题数", preview["questionCount"])
    check(preview["sections"][0]["questions"][0]["options"] == 4, "预览带选项数",
          preview["sections"][0]["questions"][0])

    print("[8] 错误形状")
    expect_error(lambda: parse_paper("   "), "text-required", "空文本报 text-required")
    auto = build_exam({"text": ENGLISH, "title": "自动存到文档目录"})
    check(auto["outputPath"].startswith(os.path.join(work, "教师助手")), "没给路径就存到 文档/教师助手/出卷",
          auto["outputPath"])
    check(auto["outputPath"].endswith(".docx") and os.path.exists(auto["outputPath"]), "真的生成了文件")

    print("[9] 回归闸：引擎不能判自己的产物 fail，也不能吞掉老师的话")
    messy_text = """七年级英语期中测试
满分 100 分，时间 90 分钟

一、听力（20 分）
1. 听句子选图（5 分）
A. apple
B. banana

二、单项选择（30 分）
3. 下列哪个是正确的
A)甲 B)乙 C)丙 D)丁
4. 二次函数 y=x^2 的图像开口向上。
5. 第三步：核对答案。
"""
    messy_path = os.path.join(work, "messy.docx")
    messy = build_exam({"text": messy_text, "outputPath": messy_path})
    document = docx.Document(messy_path)
    text = "\n".join(paragraph.text for paragraph in document.paragraphs)
    check(messy["questionCount"] == 4, "正文句子没被当大题吞掉", messy["questionCount"])
    check("二次函数" in text and "第三步" in text, "老师的话都还在卷面上", text[:120])
    check(messy["preflight"]["level"] != "fail", "自产卷不判 fail",
          [item["code"] for item in messy["preflight"]["findings"]])
    check(any(item["code"] == "section-score-mismatch" for item in messy["warnings"]),
          "大题分和小题分对不上要提示", messy["warnings"])
    check(not any(item["code"] == "placeholders-left" for item in messy["preflight"]["findings"]),
          "没有漏填的占位符")
    heading_lines = [paragraph.text for paragraph in document.paragraphs if paragraph.text.startswith("一、")]
    check(any("听力" in line for line in heading_lines), "大题序号没丢（一、听力）", heading_lines[:3])

    print("[10] 大题行里夹着第一题：第一题不能消失")
    inline = parse_paper("一、选择题1. 下列字音正确的是（ ）\nA. zhān\nB. zhǎn")
    numbers = [question["number"] for question in inline["sections"][0]["questions"]]
    check(numbers == [1], "第一题被切出来了", numbers)
    check("下列字音" in inline["sections"][0]["questions"][0]["stem"], "题干完整",
          inline["sections"][0]["questions"][0]["stem"])

    print("[11] 真实卷面写法：每小题分、答案带解析、小节标题、答题卡编号")
    detailed = parse_paper("""九年级英语模拟
一、听力（共 3 小题，每小题 1 分，满分 3 分）
1. 听句子选图
2. 听问句选答语
3. 听对话填词

二、单项选择（共 2 小题，每小题 2 分，满分 4 分）
4. 选正确的
【答案】A
【解析】因为 B 不对
""")
    first = detailed["sections"][1]
    check([question["score"] for question in first["questions"]] == [1.0, 1.0, 1.0],
          "「每小题 1 分」逐题算分，不再默认 2 分",
          [question["score"] for question in first["questions"]])
    check(first["title"] == "听力", "大题标题不带分值说明", first["title"])
    second = detailed["sections"][2]
    check([question["score"] for question in second["questions"]] == [2.0], "第二大题逐题 2 分",
          [question["score"] for question in second["questions"]])
    mixed = parse_paper("""一、选择（10 分）
1. 下列正确的是
A. 甲
B. 乙
答案：A 解析：因为乙不对
""")
    question = mixed["sections"][-1]["questions"][0]
    check(question["answer"] == "A", "答案只存答案", question["answer"])
    check("乙不对" in (question["analysis"] or ""), "解析单独存，不并进答案", question["analysis"])

    subhead = parse_paper("""八年级语文期中
一、积累与运用（20 分）
1. 字音（2 分）
2. 默写（2 分）
（一）文言文阅读（10 分）
1. 解词（2 分）
2. 翻译（3 分）
""")
    titles = [section["title"] for section in subhead["sections"]]
    check("文言文阅读" in titles, "「（一）文言文阅读」切成独立小节", titles)
    sub = [section for section in subhead["sections"] if section["title"] == "文言文阅读"][0]
    check([question["number"] for question in sub["questions"]] == [1, 2], "小节里的题号从 1 起",
          [question["number"] for question in sub["questions"]])

    card_path = os.path.join(work, "card.docx")
    card = build_exam({"text": """八年级语文期中
一、积累与运用（4 分）
1. 字音（2 分）
2. 默写（2 分）
（一）文言文阅读（4 分）
1. 解词（2 分）
2. 翻译（2 分）
""", "outputPath": card_path, "options": {"answerCard": "card", "answerSheet": "none"}})
    document = docx.Document(card_path)
    header_row = []
    for table in document.tables:
        first_col = [row.cells[0].text.strip() for row in table.rows]
        if len(table.columns) == 2 and first_col and all(value.isdigit() for value in first_col):
            header_row.extend(first_col)  # 填空/简答的「题号 + 作答区」表
        else:
            header_row.extend(cell.text for cell in table.rows[0].cells if cell.text.isdigit())
    check(header_row == ["1", "2", "3", "4"], "答题卡题号连续（小节也接着数）", header_row)
    footer_xml = document.sections[0].footer.paragraphs[0]._p.xml
    # 混了答题卡（两种纸型）时正文页脚只印当前页：两种纸型的页数加一起对谁都不对
    check("PAGE" in footer_xml and "NUMPAGES" not in footer_xml,
          "混纸型时正文页脚只印当前页码", footer_xml[:160])
    card_section = document.sections[-1]  # 8 开两栏时：标题节 / 正文节 / 答题卡节
    card_footer_xml = card_section.footer.paragraphs[0]._p.xml
    check("PAGE" in card_footer_xml and "NUMPAGES" not in card_footer_xml,
          "答题卡页脚有页码域、不印会算错的总页数", card_footer_xml[:160])
    check("w:start=\"1\"" in card_section._sectPr.xml, "答题卡页码从 1 重新开始",
          card_section._sectPr.xml[-320:])

    # 单纸型（不带答题卡/答案页）时总页数是准的，必须保留「共 N 页」
    plain_path = os.path.join(work, "plain-footer.docx")
    build_exam({"text": "七年级英语单元测\n一、听力（4 分）\n1. q1（2 分）\nA. a  B. b\n【答案】A\n",
                "outputPath": plain_path,
                "options": {"answerCard": "none", "answerSheet": "none"}})
    plain_footer = docx.Document(plain_path).sections[0].footer.paragraphs[0]._p.xml
    check("PAGE" in plain_footer and "NUMPAGES" in plain_footer,
          "单纸型卷子页脚保留「共 N 页」", plain_footer[:160])
    check(card["preflight"]["level"] != "fail", "自产语文卷不判 fail",
          [item["code"] for item in card["preflight"]["findings"]])

    print("[12] 卷尾整段答案（【答案】1—5 BCBDA…）当成答案区，不印进学生卷（persona P0）")
    tailed = """七年级英语单元测
一、听力（4 分）
1. 听句子选图（2 分）
2. 听问句选答语（2 分）

二、单项选择（4 分）
3. 选正确的（2 分）
4. 选正确的（2 分）

【答案】
1—2 AB
3—4 CD
"""
    parsed_tail = parse_paper(tailed)
    tail_numbers = [question["number"]
                    for section in parsed_tail["sections"] for question in section["questions"]]
    check(tail_numbers == [1, 2, 3, 4], "卷尾整段答案不会被当成题", tail_numbers)
    check([question["answer"] for section in parsed_tail["sections"]
           for question in section["questions"]] == ["A", "B", "C", "D"],
          "答案按题号回填到每道题",
          [question["answer"] for section in parsed_tail["sections"]
           for question in section["questions"]])
    check("answer-block" in [item["code"] for item in parsed_tail["warnings"]],
          "会告诉老师那段答案被划成答案区了",
          [item["code"] for item in parsed_tail["warnings"]])
    tail_path = os.path.join(work, "tail.docx")
    tailed_result = build_exam({"text": tailed, "outputPath": tail_path,
                                "options": {"answerCard": "none", "answerSheet": "end"}})
    check(tailed_result["questionCount"] == 4, "学生卷只有 4 题（答案不作题）",
          tailed_result["questionCount"])
    student_text = "\n".join(paragraph.text for paragraph in docx.Document(tail_path).paragraphs)
    check("【答案】" not in student_text, "学生卷上没有【答案】字样", student_text[-120:])

    print("[13] 逐题「【答案】A」不能被当成整段答案切走（回归闸）")
    inline = parse_paper("""一、听力（4 分）
1. What does he look like?
A. tall
B. short
【答案】A
2. What does she look like?
A. thin
B. fat
【答案】B
""")
    inline_questions = [question for section in inline["sections"] for question in section["questions"]]
    check(len(inline_questions) == 2, "两道题都在", len(inline_questions))
    check([question["answer"] for question in inline_questions] == ["A", "B"], "逐题答案各归各位",
          [question["answer"] for question in inline_questions])

    print("[14] 勾了答案页却一道答案都没有 → 照出卷子，但要提醒、别印空白答案页")
    noanswer = build_exam({"text": "一、听力（2 分）\n1. 听句子（2 分）\n",
                          "outputPath": os.path.join(work, "noanswer.docx"),
                          "options": {"answerCard": "none", "answerSheet": "end"}})
    noanswer_codes = [item.get("code") for item in (noanswer.get("warnings") or [])]
    check("answers-missing" in noanswer_codes, "缺答案时给 answers-missing 提醒", noanswer_codes)
    check(os.path.exists(noanswer["outputPath"]), "卷子照常出（不能因为没答案就不印）")
    noanswer_doc = docx.Document(noanswer["outputPath"])
    check(not any("参考答案" in paragraph.text for paragraph in noanswer_doc.paragraphs),
          "没答案就不印空白的参考答案页")

    print("[15] A/B 两套卷：B 卷选项顺序打乱、答案跟着改（同一份题、同一套排版）")
    ab_payload = {
        "text": """一、单项选择（共 3 小题，每小题 2 分，满分 6 分）
1. Can you swim?
A. Yes, I can.
B. No, I can't.
C. I like it.
【答案】A
2. What time is it?
A. It's nine.
B. It's Monday.
C. It's sunny.
【答案】A
3. Where is the bank?
A. Near the park.
B. At six.
C. By bus.
【答案】A
""",
        "title": "AB 卷",
        "outputPath": os.path.join(work, "ab", "卷子（A 卷）.docx"),
        "options": {"answerCard": "none", "answerSheet": "end"}
    }
    os.makedirs(os.path.join(work, "ab"), exist_ok=True)
    pair = TABLE["exam:build-ab"](ab_payload)
    check([variant["label"] for variant in pair["variants"]] == ["A", "B"], "出了 A、B 两套",
          [variant["label"] for variant in pair["variants"]])
    check(all(os.path.exists(variant["outputPath"]) for variant in pair["variants"]), "两套文件都落盘")
    check(all(variant["totalScore"] == 6.0 for variant in pair["variants"]), "两套分值一致")
    text_a = [paragraph.text for paragraph in docx.Document(pair["variants"][0]["outputPath"]).paragraphs]
    text_b = [paragraph.text for paragraph in docx.Document(pair["variants"][1]["outputPath"]).paragraphs]
    options_a = [line for line in text_a if line.startswith("A.") and "B." in line]
    options_b = [line for line in text_b if line.startswith("A.") and "B." in line]
    check(bool(options_a) and options_a != options_b, "B 卷选项顺序和 A 卷不同", [options_a[:1], options_b[:1]])
    # B 卷第 1 题的答案字母必须指向和 A 卷同一个选项文本
    def option_text(lines, question_index, letter):
        block = [line for line in lines if line.startswith("A.") and "B." in line][question_index]
        for piece in block.split("    "):
            if piece.strip().startswith(letter + "."):
                return piece.split(".", 1)[1].strip()
        return None
    answer_a = [line for line in text_a if line.strip().startswith("1. ") and "Can you swim" in line][0]
    check(option_text(text_b, 0, "B") is not None, "B 卷答案字母能对上选项")
    b_options = [line for line in text_b if line.startswith("A.") and "B." in line]
    check("Yes, I can." in b_options[0], "B 卷第 1 题里仍有正确选项（文本没丢）", b_options[0])

    print("[16] 答题卡给作文题留作答格（persona：书面表达只有标题没有格）")
    essay_path = os.path.join(work, "essay-card.docx")
    essay = build_exam({"text": """七年级英语期末
一、单项选择（4 分）
1. — Can you swim?
A. Yes, I can.  B. No, I can't.
2. — What time is it?
A. It's nine.  B. It's Monday.
二、书面表达（10 分）
3. 假设你是李华，给你的英国朋友写一封邮件介绍你的学校，不少于 60 词。
""", "outputPath": essay_path,
        "options": {"answerCard": "card", "answerSheet": "none"}})
    essay_doc = docx.Document(essay_path)
    # 答题卡正文段：作文作答提示 + 作答格（默认 10 行单列表格）
    card_texts = [paragraph.text for paragraph in essay_doc.paragraphs]
    check(any("请在下面的格子里作答" in line for line in card_texts),
          "答题卡上有作文作答提示", card_texts[-2:])
    # 卡上表格：两块选择题表（各 2 列）+ 一张 10 行单列作答格
    card_tables = [table for table in essay_doc.tables if len(table.columns) == 1]
    check(bool(card_tables) and len(card_tables[0].rows) == 10,
          "作文作答格是 10 行单列表格",
          [len(t.rows) for t in card_tables])
    first_cell_height = card_tables[0].rows[0].height
    check(first_cell_height is not None and first_cell_height.mm >= 8,
          "作答格行高够写字", str(first_cell_height))
    check(not any(cell.text.isdigit() for row in card_tables[0].rows for cell in row.cells),
          "作答格里没有题号数字（学生不往里写字母）")
    check(any("班级" in line and "姓名" in line for line in card_texts),
          "答题卡抬头有班级/姓名/学号（persona：没地方写名字）", card_texts[:4])
    check(essay["preflight"]["level"] != "fail", "带作文格的卷子不判 fail",
          [item["code"] for item in essay["preflight"]["findings"]])

    print("[17] 学生卷 + 教师卷一次出两份（persona：讲评时不用翻页找答案）")
    pair_teacher = TABLE["exam:build-teacher"]({
        "text": """七年级英语练习
一、单项选择（4 分）
1. — Can you swim?
A. Yes, I can.  B. No, I can't.
【答案】A
【解析】can 引导的一般疑问句用 Yes, I can. 回答。
2. — What time is it?
A. It's nine.  B. It's Monday.
【答案】A
""",
        "title": "教师卷测试",
        "outputPath": os.path.join(work, "teacher", "练习（学生卷）.docx"),
        "options": {"answerCard": "none", "answerSheet": "none"},
    })
    os.makedirs(os.path.join(work, "teacher"), exist_ok=True)
    check(os.path.exists(pair_teacher["outputPath"]) and os.path.exists(pair_teacher["teacherPath"]),
          "学生卷与教师卷都落盘", [pair_teacher["outputPath"], pair_teacher.get("teacherPath")])
    check("（教师卷）" in pair_teacher["teacherPath"], "教师卷文件名带「（教师卷）」", pair_teacher["teacherPath"])
    # 卷头（卷名/班级行）在两栏卷子里落在第一页页眉，所以取文本要连页眉一起取
    def doc_text(path):
        document = docx.Document(path)
        parts = []
        for section in document.sections:
            parts.extend(paragraph.text for paragraph in section.first_page_header.paragraphs)
        parts.extend(paragraph.text for paragraph in document.paragraphs)
        return parts

    student_text = doc_text(pair_teacher["outputPath"])
    teacher_text = doc_text(pair_teacher["teacherPath"])
    check(not any("【答案】" in line for line in student_text), "学生卷上没有答案", student_text[:4])
    check(sum(1 for line in teacher_text if "【答案】" in line) == 2, "教师卷每道题后都有答案",
          [line for line in teacher_text if "【答案】" in line])
    check(any("【解析】" in line for line in teacher_text), "教师卷带解析")
    check(pair_teacher["teacher"]["questionCount"] == pair_teacher["questionCount"] == 2,
          "两份卷子题目数一样")
    check(any("（教师卷）" in line for line in teacher_text), "教师卷标题上标了「（教师卷）」", teacher_text[:2])
    check("姓名" not in "\n".join(teacher_text), "教师卷不印班级/姓名/学号那行（自己讲评用）",
          [line for line in teacher_text if "姓名" in line])
    check(any("姓名" in line for line in student_text), "学生卷仍印班级/姓名/学号",
          [line for line in student_text if "姓名" in line])

    print("[18] 阅读理解的材料段（B / 材料：）不许塞进上一题题干（persona P0）")
    material_path = os.path.join(work, "material.docx")
    material = build_exam({"text": """七年级英语阅读专练
一、阅读理解（6 分）
1. What is Tom's favourite sport?
A. Football  B. Basketball
2. How old is Tom?
A. Ten  B. Twelve
B
Tom is a 12-year-old boy from London. He likes football very much.
3. What does Tom often do after school?
A. Play football  B. Watch TV
材料：
Mary is a Chinese girl. She studies in a middle school in Beijing.
4. Where does Mary study?
A. In London  B. In Beijing
""", "outputPath": material_path,
        "options": {"answerCard": "none", "answerSheet": "none"}})
    texts = [paragraph.text for paragraph in docx.Document(material_path).paragraphs]
    check(material["questionCount"] == 4, "四道题都在（材料没被当题）",
          material.get("questionCount"))
    check(any("Tom is a 12-year-old boy" in line for line in texts), "B 篇短文印出来了",
          [line[:24] for line in texts if "Tom" in line])
    check(any("Mary is a Chinese girl" in line for line in texts), "「材料：」开头的短文也印出来了",
          [line[:24] for line in texts if "Mary" in line])
    stem_of_2 = next((line for line in texts if line.startswith("2.")), "")
    check("Tom is a 12-year-old" not in stem_of_2, "第 2 题的题干里没有整篇短文", stem_of_2[:60])
    order = [index for index, line in enumerate(texts) if "Mary" in line or line.startswith("4.")]
    check(len(order) == 2 and order[0] < order[1], "Mary 那篇短文排在第 4 题前面", order)
    check(material["preflight"]["level"] != "fail", "带材料的卷子不判 fail",
          [item["code"] for item in material["preflight"]["findings"]])

    print("[19] 页脚印真班级/真老师（不填就留手写横线）")
    named_path = os.path.join(work, "named.docx")
    named = build_exam({"text": """七年级英语小测
一、单项选择（2 分）
1. — Can you swim?
A. Yes, I can.  B. No, I can't.
【答案】A
""", "outputPath": named_path,
        "values": {"className": "七(3)班", "teacherName": "王老师"},
        "options": {"answerCard": "none", "answerSheet": "none"}})
    footer_text = docx.Document(named_path).sections[0].footer.paragraphs[0].text
    check("七(3)班" in footer_text, "页脚印了使用班级", footer_text)
    check("王老师" in footer_text, "页脚印了出卷老师", footer_text)
    check("__________" not in footer_text, "填了真名就不再留横线", footer_text)
    check(named["values"]["className"] == "七(3)班", "values 回传面板", named["values"])
    plain_path = os.path.join(work, "plain-footer-2.docx")
    build_exam({"text": """七年级英语小测
一、单项选择（2 分）
1. — Can you swim?
A. Yes, I can.  B. No, I can't.
""", "outputPath": plain_path,
        "options": {"answerCard": "none", "answerSheet": "none"}})
    plain_footer = docx.Document(plain_path).sections[0].footer.paragraphs[0].text
    check("__________" in plain_footer and "{{" not in plain_footer,
          "没填班级/老师就留手写横线（不印占位符）", plain_footer)

    print("[20] 答题卡给填空/简答题留作答格（只给题号格学生没处写字）")
    fill_path = os.path.join(work, "fillcard.docx")
    filled = build_exam({"text": """七年级英语小测
一、单项选择（4 分）
1. — Can you swim?
A. Yes, please.  B. No, thanks.
【答案】A
2. This is ______ apple.
A. a  B. an
【答案】B
二、用所给词的正确形式填空（4 分）
3. He ______ (be) tall.
4. She ______ (have) a bag.
""", "outputPath": fill_path, "options": {"answerCard": "card", "answerSheet": "end"}})
    card_doc = docx.Document(fill_path)
    number_slots = []
    slot_table = None
    for table in card_doc.tables:
        first = [row.cells[0].text.strip() for row in table.rows]
        if len(table.columns) == 2 and first and all(value.isdigit() for value in first):
            number_slots.append(first)
            if first == ["3", "4"]:
                slot_table = table
    check(bool(number_slots), "答题卡上有一张「题号 + 作答区」的填空表格", len(card_doc.tables))
    check(number_slots[:1] == [["3", "4"]], "填空题的题号接着选择题往下排", number_slots)
    check(slot_table is not None and slot_table.rows[0].height is not None,
          "作答格有行高（手写留够地方）",
          None if slot_table is None else round(slot_table.rows[0].height.mm, 1))
    check(filled["preflight"]["level"] != "fail", "带填空作答格的卷子不判 fail",
          [item["code"] for item in filled["preflight"]["findings"]])

    print("[21] 完形填空走字母题号格，教师卷答案跟着题走，模板纸型不符要提醒")
    # persona P2：完形 11—15 在答题卡上被排成「题号+作答区」的竖排格（答案明明是字母），
    # preflight 还会因为两张表题号不连续误报
    cloze_path = os.path.join(work, "clozecard.docx")
    build_exam({"text": """七年级英语练习
一、完形填空（共 3 小题，每小题 2 分，共 6 分）
11. He ___ tall.
A. is  B. are  C. am
【答案】A
12. She ___ a bag.
A. have  B. has  C. having
【答案】B
13. They ___ here.
A. is  B. am  C. are
【答案】C
二、书面表达（10 分）
14. 写一段话介绍你的朋友。
【答案】略
""", "outputPath": cloze_path, "options": {"answerCard": "card", "answerSheet": "end"}})
    cloze_doc = docx.Document(cloze_path)
    letter_rows = []
    writing_tables = []
    for table in cloze_doc.tables:
        header = [cell.text.strip() for cell in table.rows[0].cells]
        column0 = [row.cells[0].text.strip() for row in table.rows]
        if header and header[0] == "题号":
            letter_rows.extend(cell for cell in header[1:] if cell)
        elif len(table.columns) == 2 and column0 and all(value.isdigit() for value in column0):
            writing_tables.append(column0)
    check(letter_rows == ["1", "2", "3"],
          "完形填空给的是字母题号格（不再竖排作答区）", letter_rows)
    check(not writing_tables, "完形填空没有生成书写作答格", writing_tables)

    # 教师卷：答案必须和上一段（最后一个选项）拴在一起，不能单独浮到右栏顶上
    teacher_path = os.path.join(work, "keepanswer.docx")
    build_exam({"text": """七年级英语小测
一、单项选择（4 分）
1. — Can you swim?
A. Yes, please.  B. No, thanks.
【答案】A
2. This is ______ apple.
A. a  B. an
【答案】B
""", "outputPath": teacher_path, "options": {"inlineAnswers": True}})
    teacher_doc = docx.Document(teacher_path)
    paragraphs = teacher_doc.paragraphs
    bound = 0
    for index, paragraph in enumerate(paragraphs):
        if paragraph.text.strip().startswith("【答案】") and index > 0:
            if paragraphs[index - 1].paragraph_format.keep_with_next:
                bound += 1
    check(bound == 2, "教师卷每道题的答案都和上一段拴在一起", bound)

    # 模板纸型与手选纸型不一致：要提醒（老板会以为"套了模板就是学校那份"）
    template_path = os.path.join(work, "16k-template.docx")
    TABLE["docx:new"]({"outputPath": template_path, "page": {"size": "16K", "orientation": "portrait", "columns": 1},
                       "blocks": [{"type": "paragraph", "text": "{{examTitle}}"}]})
    mismatch_path = os.path.join(work, "mismatch.docx")
    mismatch = build_exam({"text": """七年级英语小测
1. This is ______ apple.
A. a  B. an
【答案】B
""", "outputPath": mismatch_path, "baseTemplate": template_path,
        "page": {"size": "8K", "orientation": "landscape", "columns": 2}})
    codes = [item["code"] for item in mismatch["warnings"]]
    check("template-paper-mismatch" in codes, "模板纸型和手选纸型不一致时给提醒", codes)
    mismatch_doc = docx.Document(mismatch_path)
    check(round(mismatch_doc.sections[0].page_width.mm) == 368, "提醒之后仍按老师这次选的 8 开排",
          round(mismatch_doc.sections[0].page_width.mm))

    print("[21b] 模板页脚带页码时不印错的总页数")
    footer_template = os.path.join(work, "footer-template.docx")
    TABLE["docx:new"]({"outputPath": footer_template, "page": {"size": "16K", "orientation": "portrait", "columns": 1},
                       "blocks": [{"type": "paragraph", "text": "{{examTitle}}"}]})
    footer_doc = docx.Document(footer_template)
    paragraph = footer_doc.sections[0].footer.paragraphs[0]
    paragraph.add_run("第")
    _field = OxmlElement("w:fldSimple")
    _field.set(qn("w:instr"), " PAGE ")
    _run = OxmlElement("w:r")
    _node = OxmlElement("w:t")
    _node.text = "1"
    _run.append(_node)
    _field.append(_run)
    paragraph._p.append(_field)
    paragraph.add_run("页 共")
    _field2 = OxmlElement("w:fldSimple")
    _field2.set(qn("w:instr"), " NUMPAGES ")
    _run2 = OxmlElement("w:r")
    _node2 = OxmlElement("w:t")
    _node2.text = "1"
    _run2.append(_node2)
    _field2.append(_run2)
    paragraph._p.append(_field2)
    paragraph.add_run("页")
    footer_doc.save(footer_template)
    tpl_multi = os.path.join(work, "tpl-multi.docx")
    TABLE["exam:build"]({"text": """七年级英语小测
1. This is ______ apple.
A. a  B. an
【答案】B
""", "outputPath": tpl_multi, "baseTemplate": footer_template,
        "options": {"answerCard": "A4", "answerSheet": "pages"}})
    tpl_multi_doc = docx.Document(tpl_multi)
    foot_text = "".join(node.text or "" for node in tpl_multi_doc.sections[0].footer._element.iter(qn("w:t")))
    check("共" not in foot_text, "多节卷子的模板页脚不再印算错的总页数", foot_text)
    check("第" in foot_text and "页" in foot_text, "页码本身还在", foot_text)
    check(len(tpl_multi_doc.sections) >= 2, "确实分出了答题卡/答案页", len(tpl_multi_doc.sections))

    print("[22] A/B 两套 + 两份教师卷（四件一起出）")
    # 老师既想要 B 卷，又想把 A 卷的答案讲给学生 —— 以前只能在面板里二选一
    four_path = os.path.join(work, "four.docx")
    four = TABLE["exam:build-ab"]({
        "text": """七年级英语四份
一、单项选择（4 分）
1. — Can you swim?
A. Yes, please.  B. No, thanks.
【答案】A
2. This is ______ apple.
A. a  B. an
【答案】B
""",
        "outputPath": four_path,
        "options": {"teacher": True, "answerCard": "card", "answerSheet": "end"},
    })
    labels = [item["label"] for item in four["variants"]]
    check(labels == ["A", "B"], "出的是 A、B 两套", labels)
    check(len(four.get("teacherPaths") or []) == 2, "另外给了两份教师卷的路径", four.get("teacherPaths"))
    from_students = [item["outputPath"] for item in four["variants"]]
    from_teachers = list(four.get("teacherPaths") or [])
    check(all(os.path.exists(path) for path in from_students + from_teachers),
          "四份文件都落盘了", from_students + from_teachers)
    check(from_students[0] == four_path and from_students[1].endswith("-B.docx"),
          "老师指定了输出路径时，B 卷只在旁边加 -B（不擅自改名）",
          [os.path.basename(path) for path in from_students])
    default_named = TABLE["exam:build-ab"]({"text": "七年级英语四份\n1. This is ______ apple.\nA. a  B. an\n【答案】B\n",
                                            "options": {"teacher": True}})
    student_names = [os.path.basename(item["outputPath"]) for item in default_named["variants"]]
    check(all(("（A 卷）" in name or "（B 卷）" in name) for name in student_names),
          "没指定路径时学生卷按「（A 卷）/（B 卷）」命名", student_names)
    teacher_names = [os.path.basename(path) for path in (default_named.get("teacherPaths") or [])]
    check(len(teacher_names) == 2 and all("（教师卷）" in name for name in teacher_names),
          "没指定路径时教师卷带着「（教师卷）」", teacher_names)
    check(len(set(path.lower() for path in from_students + from_teachers)) == 4,
          "四份分别是四个文件，没有互相覆盖", from_students + from_teachers)
    teacher_texts = []
    for path in from_teachers:
        document = docx.Document(path)
        joined = "\n".join(paragraph.text for paragraph in document.paragraphs)
        teacher_texts.append(joined)
    check(all("【答案】" in text for text in teacher_texts), "两份教师卷每题下面都印了答案",
          [text.count("【答案】") for text in teacher_texts])
    student_texts = []
    for path in from_students:
        document = docx.Document(path)
        student_texts.append("\n".join(paragraph.text for paragraph in document.paragraphs))
    check(all("【答案】" not in text for text in student_texts), "学生卷里没有答案",
          [text.count("【答案】") for text in student_texts])
    check(len({round(item["totalScore"], 2) for item in four["variants"]}
              | {round(value, 2) for value in [four["totalScore"]]}) == 1,
          "四份卷子的满分一致", four["totalScore"])
    check(len(four.get("warnings") or []) == 0, "四件一起出不产生额外提醒", four.get("warnings"))

    print("[23] 作文大题只有一段要求、没有题号（persona：二、书面表达 整个标题消失）")
    # 以前的毛病：没有题号的节被并进上一节 → 「二、书面表达」标题不见了，
    # 作文要求印在「一、单项选择」标题下面、第 1 题前面，答题卡上也没有作文格
    writing_path = os.path.join(work, "writing-nonum.docx")
    writing = build_exam({
        "text": """七年级英语阶段测试卷
一、单项选择（每题 2 分，共 4 分）
1. — How ___ your weekend?
A. is  B. was
2. She ___ to the library yesterday.
A. go  B. went
二、书面表达（10 分）
请以 My Weekend 为题写一篇 60 词左右的短文。
""",
        "outputPath": writing_path,
        "options": {"answerCard": "card", "answerSheet": "none"},
    })
    writing_doc = docx.Document(writing_path)
    writing_texts = [paragraph.text for paragraph in writing_doc.paragraphs]
    writing_sections = [item for item in writing["sections"] if item.get("writing")]
    check(len(writing_sections) == 1 and writing_sections[0]["title"] == "书面表达",
          "没有题号的作文大题单独成节，标题还在", writing["sections"])
    check(any("二、书面表达" in line for line in writing_texts),
          "卷面上印出「二、书面表达」标题", writing_texts)
    essay_index = next(index for index, line in enumerate(writing_texts) if "二、书面表达" in line)
    prompt_index = next(index for index, line in enumerate(writing_texts) if "My Weekend" in line)
    check(prompt_index > essay_index, "作文要求排在它自己的标题下面", writing_texts[essay_index:prompt_index + 1])
    first_question = next(index for index, line in enumerate(writing_texts) if line.startswith("1. "))
    check(essay_index > first_question, "作文节排在选择题后面", [essay_index, first_question])
    check(abs(float(writing["totalScore"]) - 14.0) < 0.01,
          "作文标题里写的 10 分算进满分（满分 4 + 10 = 14）", writing["totalScore"])
    writing_tables = [table for table in writing_doc.tables if len(table.columns) == 1]
    check(bool(writing_tables) and len(writing_tables[0].rows) == 10,
          "答题卡上给没有题号的作文留了 10 行作答格", [len(t.rows) for t in writing_tables])

    print()
    if FAILS:
        print("失败 " + str(len(FAILS)) + " 项：")
        for label in FAILS:
            print(" - " + label)
        sys.exit(1)
    print("全部通过 ✅")


if __name__ == "__main__":
    main()
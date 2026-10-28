"""听写/默写单：老师粘一列「英文 中文」，出一张能直接印的 A4 纸。

为什么单独一个模块，不塞进 exam.py：
- 听写是高频轻活（一周 2—3 次），老师的痛点是"在 Word 里排半小时"。所以默认值
  必须一次就对（A4 竖版、2 栏、2 倍行距、14 个下划线），老师「粘进去、点一下、打印」，
  中间不该再调参数。
- 版式和试卷相反：试卷的题干要跨栏、要按题号分节；听写单只要求一行一条、条目在栏里
  自然流动。正文绝不用表格排——表格在双栏里会被压窄或跨栏（见
  docx_build.column_width_twips 的注释），印出来最扎眼。
- 输入是脏文本：粘贴常带 1. 序号、Tab、全角逗号、空格数不一。解析的容错比排版更
  重要——少一条学生就没得写，所以拆不开的条目宁可保留并给警告，也不静默丢弃。

两种工作方式（mode）只是"谁露脸"不同，原始数据始终是「英文 + 中文」一对：
en2cn 印英文让学生写中文；cn2en 反过来；mix 按奇偶交替，防同桌互抄。
"""

import io
import os
import re

from docx import Document
from docx.shared import Pt
from docx.table import Table
from docx.text.paragraph import Paragraph

from .docx_build import build_docx
from .errors import EngineError, require
from .paths import default_output

MAX_ITEMS = 120          # 一次一份最多 120 条：再多，A4 双栏也得印五六页，学生写不完
MAX_COLUMNS = 3
DEFAULT_COLUMNS = 2      # A4 一栏太宽（一行跑半页），三栏每栏只剩 44mm，写中文挤
DEFAULT_BLANK_LENGTH = 14
DEFAULT_LINE_SPACING = 2.0   # 2 倍行距＝写完一行还留得下手写的高度
DEFAULT_TITLE = "英语听写"
ANSWER_HEADING = "参考答案"
INFO_LINE = "班级：__________　姓名：__________　得分：______"
MODES = ("en2cn", "cn2en", "mix")
MODE_HINT = "可用：en2cn（给英文，写中文）、cn2en（给中文，写英文）、mix（奇偶交替，防抄）"
TEXT_HINT = "每行一条：英文和中文用空格或 Tab 隔开"

CJK_RANGE = "\u3400-\u4dbf\u4e00-\u9fff\uf900-\ufaff"
RE_CJK = re.compile("[" + CJK_RANGE + "]")
RE_LATIN = re.compile(r"[A-Za-z]")

# 行首序号：老师的词表几乎都带号（1. / 1、/ 1) / （1）），带号不剥会和我们的重排撞车
RE_NUMBERING = re.compile(r"^\s*(?:[(（]\s*\d{1,3}\s*[)）]|\d{1,3}\s*[.、)）．])\s*")
RE_TWO_SPACES = re.compile(r"[ \u3000]{2,}")
# 中英交界：一个空格也算（"apple 苹果"），紧贴着不空格也算（"be good at擅长"）
RE_CJK_BOUNDARY = re.compile(
    r"(?<=[A-Za-z0-9])[ \u3000]*(?=[" + CJK_RANGE + r"])"
    r"|(?<=[" + CJK_RANGE + r"])[ \u3000]*(?=[A-Za-z0-9])"
)
RE_COMMA_CN = re.compile(r"[，、]")
RE_COMMA_ASCII = re.compile(r",[ \u3000]")   # "apple, 苹果"；光秃秃的 "," 不拆（"1,000"）
RE_EQUAL = re.compile(r"[=＝]")
RE_DASH = re.compile(r"[—–]|-{2,}")
RE_HYPHEN = re.compile(r"[ \u3000]+-[ \u3000]+")   # 带空格的 "-" 才拆：ice-cream 不能被切断
RE_COLON = re.compile(r"[：:]")
RE_BRACKET_EN = re.compile(r"^(.+?)[（(][ \u3000]*([^（()）]+?)[ \u3000]*[)）]$")
RE_ITEM_PARA = re.compile(r"^\d{1,3}[.．][ \u3000]")
# 老师粘词表时第一行常是标题（"Unit 2 词汇听写（20 个）"）：没有分隔符、带听写/单元字样，
# 当条目收进去会变成"1. Unit 2 ______"，答案页还印着"-- 词汇听写（20 个）"
RE_HEADING = re.compile(r"(听写|默写|小测|练习|测试|词汇表|单词表)", re.I)
# 抬头/落款里才会出现的词：老师的配套词表是「首段标题 + 两列表格 + 末段落款」，
# 标题常常是一整句（"Unit 3 What does he look like? 单词表"），按词数卡不住，只能按词认。
# 落款更不会自己说"我是落款"——「姓名/班级/日期/签名/admin」这些词就是它的指纹。
RE_HEADING_WIDE = re.compile(
    r"(单词表|词汇表|词表|听写单|听写|默写|小测|单元|Unit\s*[0-9一二三四五六七八九十]"
    r"|姓名|班级|日期|签名)", re.I)
# 句子感：真词条的英文半边是单词或短语，不会带数字、问号、逗号这些句子零件
RE_SENTENCE_MARK = re.compile(r"[0-9?？!！,，;；。]")

SPACE_JUNK = " \u3000\t"
# 表格里的辅助列：序号（"1" / "1." / "①"）、音标（/…/ 或含 IPA 专用符号）、词性（adj.）
RE_INDEX_CELL = re.compile(r"^[（(]?\d{1,3}[)）.．、]?$")
RE_IPA = re.compile(r"[ˈˌəɔʃʒŋθðʌɑɒɪʊæɜː]")
RE_POS_CELL = re.compile(r"^(?:n|v|vt|vi|adj|adv|prep|conj|pron|num|art|int|aux)\.?$", re.I)


# ---------------------------------------------------------------- 解析

def _norm_mode(mode):
    """方式名不合法要当场报，不能默默当成 en2cn——印错了整节课就废了。"""
    value = (mode if mode is not None else "en2cn")
    value = str(value).strip().lower()
    if value not in MODES:
        raise EngineError("bad-mode", "不认识的听写方式：" + str(mode), MODE_HINT)
    return value


def _item_mode(mode, index):
    """单条的分向：mix 用 1 基序号判奇偶（第 1 条给英文、第 2 条给中文……）。"""
    if mode != "mix":
        return mode
    return "en2cn" if index % 2 == 1 else "cn2en"


def _first_split(text, pattern):
    """按 pattern 在 text 上切第一刀；切出来的半边为空（分隔符贴在行首行尾）就当没切。"""
    for match in pattern.finditer(text):
        left = text[:match.start()].strip(SPACE_JUNK)
        right = text[match.end():].strip(SPACE_JUNK)
        if left and right:
            return left, right
        return None
    return None


def _split_line(line):
    """一个条目 → (英文, 中文)。

    分隔符按"老师真实粘进来的可能性"排优先级：Tab 最可信（Excel / 背单词软件导出），
    其次两个以上空格，然后才是中文标点。顺序就是判据：
    "hello world 你好" 必须按空格数切，不能先按任何标点切；
    而 "ice-cream 冰淇淋" 里不带空格的 "-" 是单词的一部分，不能切。
    """
    text = line.strip(SPACE_JUNK)
    if not text:
        return "", ""

    # 1) Tab（两边可能不止两个字段，右边多余的部分并回去）
    if "\t" in text:
        head, _, tail = text.partition("\t")
        tail = " ".join(piece.strip() for piece in tail.split("\t") if piece.strip())
        if head.strip(SPACE_JUNK) and tail:
            return head.strip(SPACE_JUNK), tail

    # 2) 两个以上空格 / 全角空格
    pair = _first_split(text, RE_TWO_SPACES)
    if pair:
        return pair

    # 3) 中英交界（一个空格，或者干脆没空格）
    pair = _first_split(text, RE_CJK_BOUNDARY)
    if pair:
        return pair

    # 4) 中文逗号 / 顿号 / 英文逗号加空格
    for pattern in (RE_COMMA_CN, RE_COMMA_ASCII):
        pair = _first_split(text, pattern)
        if pair:
            return pair

    # 5) = ＝
    pair = _first_split(text, RE_EQUAL)
    if pair:
        return pair

    # 6) 破折号 —— — –，以及 --
    pair = _first_split(text, RE_DASH)
    if pair:
        return pair

    # 7) 两侧带空格的连字符
    pair = _first_split(text, RE_HYPHEN)
    if pair:
        return pair

    # 8) 冒号：只有右边确实是中文才认（"Unit 1: apple" 是标题，不能拆成条目）
    for match in RE_COLON.finditer(text):
        left = text[:match.start()].strip(SPACE_JUNK)
        right = text[match.end():].strip(SPACE_JUNK)
        if left and right and RE_CJK.search(right):
            return left, right

    # 9) 括号里是中文："apple（苹果）"。英文的 "(at)" 不算——括号里没中文就不是释义
    bracket = RE_BRACKET_EN.match(text)
    if bracket:
        inner = bracket.group(2).strip(SPACE_JUNK)
        if RE_CJK.search(inner):
            return bracket.group(1).strip(SPACE_JUNK), inner

    return text.strip(" \u3000\t，,、；;=＝:："), ""


def _looks_chinese(text):
    return bool(RE_CJK.search(text or ""))


def _has_latin(text):
    return bool(RE_LATIN.search(text or ""))


def _dedupe_key(pair, mode):
    """去重用哪一边：老师眼睛对的是印在卷面上的那一边，所以按 prompt 去重。

    mix 例外——同一条在不同奇偶位上 prompt 不同，只能按「英文+中文」整对去重，
    否则 "apple 苹果" 和 "apple 苹果的" 会被一条条错杀。
    """
    english, chinese = pair
    if mode == "cn2en":
        return ("cn", chinese or english)
    if mode == "mix":
        return ("mix", english, chinese)
    return ("en", english or chinese)


def _heading_like(body, left, right):
    """这一行像不像"词表标题"而不是第 1 个词条。

    "Unit 2 词汇听写（20 个）" / "Unit 3 单词听写" 是标题；
    "dictation 听写" 是真词条（右边就两个字）；"apple 苹果" 当然也是词条。
    """
    if len(body) > 40 or not RE_HEADING.search(body):
        return False
    if not right:
        return True                       # 没有分隔符：整行都是标题
    if len(right) > 4:
        return True                       # 右边一大串，是标题不是释义
    # 右边只剩"单词听写"这类元信息（本身就是标题词、且不止两个字），也当标题；
    # "practice 练习"、"test 测试" 右边只有两个字，是正经词条，不能跳
    if len(right) >= 4 and RE_HEADING.search(right) and not _has_latin(right):
        return True
    return False


def parse_items(text, mode="en2cn"):
    """把老师粘的一列文本拆成条目。

    返回 {"items":[{index, prompt, answer, given, blank}], "warnings":[{code, message, ...}]}

    - index：重排后的 1..N（行首原有 1. / 1、 一律剥掉，免得和学生看到的题号打架）
    - prompt：印在卷面上的那一半（学生看着它写）
    - answer：学生要写的那一半
    - given / blank：prompt / answer 的别名，给"露谁填谁"的模板层用
    """
    mode = _norm_mode(mode)
    lines = (text or "").replace("\r\n", "\n").replace("\r", "\n").split("\n")
    warnings = []
    pairs = []

    skipped_heading = 0
    skipped_lines = []
    # 先只留非空行：判"是不是最后一行"要看位置，空行不该占位置
    bodies = []
    for raw in lines:
        body = RE_NUMBERING.sub("", (raw or "").strip()).strip()
        if body:
            bodies.append(body)
    for position, body in enumerate(bodies):
        left, right = _split_line(body)
        if not left and not right:
            continue
        # 还没收到任何条目时，第一行像标题就跳过（并告诉老师），别把它当第 1 个单词。
        # 判据：含"听写/默写/小测"这类元信息，且"释义"那半边不像两三个字的词义
        # （"dictation 听写"是真词条，右边就俩字；"Unit 2 词汇听写（20 个）"不是）。
        # 末尾一行同理：老师的词表常在最后写一行落款（"（Unit 3 单词表 · 王老师）"），
        # 从 Word / txt 粘过来的整段文本不能把落款当第 N 条印出去；
        # 真词条（"test 测试"、"name 姓名"）由两段式判据保下来，不会被这里吞掉。
        heading = False
        if not pairs:
            heading = _heading_like(body, left, right) or _looks_like_heading(body, "first")
        elif position == len(bodies) - 1:
            heading = _looks_like_heading(body, "last")
        if heading:
            skipped_heading += 1
            skipped_lines.append(body[:24])
            continue
        if not left:                      # 行首就是分隔符（"，苹果"）：把有内容的一边扶正
            left, right = right, ""
        # 老师偶尔把中文写在前面（"苹果 apple"），统一成「英文在前」
        if _looks_chinese(left) and not _has_latin(left) and _has_latin(right):
            left, right = right, left

        if not right:
            ordinal = len(pairs) + 1
            if _has_latin(left):
                warnings.append({
                    "code": "no-chinese", "index": ordinal,
                    "message": "第 " + str(ordinal) + " 条只有英文，没有中文释义",
                })
            else:
                warnings.append({
                    "code": "no-english", "index": ordinal,
                    "message": "第 " + str(ordinal) + " 条只有中文，没有英文",
                })
        pairs.append((left, right))

    if skipped_heading:
        warnings.append({
            "code": "skipped-title", "count": skipped_heading,
            "message": "有 " + str(skipped_heading) + " 行像标题或落款（"
                       + " | ".join(skipped_lines) + "），我没当词条收进来",
            "lines": skipped_lines,
        })

    # 去重（保序）：同一份词表被粘了两遍是最常见的复制事故
    kept = []
    seen = set()
    duplicates = 0
    for pair in pairs:
        key = _dedupe_key(pair, mode)
        if key in seen:
            duplicates += 1
            continue
        seen.add(key)
        kept.append(pair)
    if duplicates:
        warnings.append({
            "code": "duplicate", "count": duplicates,
            "message": "有 " + str(duplicates) + " 条重复，已去掉",
        })

    if len(kept) > MAX_ITEMS:
        dropped = len(kept) - MAX_ITEMS
        warnings.append({
            "code": "too-many", "count": dropped,
            "message": "一次最多排 " + str(MAX_ITEMS) + " 条，后面 " + str(dropped) + " 条没有排进去",
        })
        kept = kept[:MAX_ITEMS]

    items = []
    for index, (english, chinese) in enumerate(kept, 1):
        if _item_mode(mode, index) == "en2cn":
            prompt, answer = english, chinese
        else:
            prompt, answer = chinese, english
        if not prompt:                    # 只有中文的条目在 en2cn 下没有 prompt，别印出空行
            prompt, answer = answer, ""
        items.append({
            "index": index,
            "prompt": prompt,
            "answer": answer,
            "given": prompt,
            "blank": answer,
        })

    return {"items": items, "warnings": warnings}


# ---------------------------------------------------------------- 排版

def _to_int(value, code, message, hint, low, high):
    try:
        number = int(float(str(value).strip()))
    except (TypeError, ValueError):
        raise EngineError(code, message, hint)
    if number < low or number > high:
        raise EngineError(code, message, hint)
    return number


def _to_float(value, code, message, hint, low, high):
    try:
        number = float(str(value).strip())
    except (TypeError, ValueError):
        raise EngineError(code, message, hint)
    if number < low or number > high:
        raise EngineError(code, message, hint)
    return number


def _item_block(item, blank_length):
    """一行一条：1. word ____________________。

    横线用下划线字符，不用表格也不用制表位——下划线跟着文字走，
    分栏换行时不会跑到别的栏里去。
    """
    text = str(item["index"]) + ". " + item["prompt"]
    if blank_length > 0:
        return {"type": "paragraph", "runs": [{"text": text + " "},
                                              {"text": "_" * blank_length}]}
    return {"type": "paragraph", "runs": [{"text": text}]}


def _answer_line(item, mode):
    """参考答案一律「英文 — 中文」，顺序不随 mode 变：老师对答案时只认这个顺序。"""
    if _item_mode(mode, item["index"]) == "cn2en":
        chinese, english = item["prompt"], item["answer"]
    else:
        english, chinese = item["prompt"], item["answer"]
    head = str(item["index"]) + ". "
    if english and chinese:
        return head + english + " — " + chinese
    return head + (english or chinese)


def _tighten_items(path, line_spacing):
    """给条目段落单独上「2 倍行距、段后 0」。

    为什么排在 build_docx 之后再改一次：docx_build 的段落块只认
    align/style/runs/keepWithNext 这些开关，没有行距字段（那里是版式层，动它等于改别人
    的公共契约）。只改条目段落而不是 Normal 样式——否则标题和说明行也被拉开，
    纸头上凭空多出半截面。
    """
    document = Document(path)
    touched = 0
    for paragraph in document.paragraphs:
        if not RE_ITEM_PARA.match(paragraph.text or ""):
            continue
        fmt = paragraph.paragraph_format
        fmt.line_spacing = float(line_spacing)
        fmt.space_before = Pt(0)
        fmt.space_after = Pt(0)
        touched += 1
    if touched:
        document.save(path)
    return touched


# ---------------------------------------------------------------- 从文件读词表

WORD_FILE_HINT = "点「浏览…」选一个 .docx / .xlsx / .csv / .txt，或者直接把词表粘进输入框"
# 表头行：「英文」「中文」「单词」「释义」「English」「Chinese」这些不是词条，读的时候直接跳过。
HEADER_WORDS = ("英文", "中文", "单词", "词汇", "词条", "释义", "意思", "翻译", "注音", "音标",
                "序号", "编号", "unit", "word", "english", "chinese", "meaning", "no")


def _row_to_entry(cells):
    """表格里的一行 → 一条词条文本（「英文 中文」）。

    老写法把整行 cells 用空格 join，于是教材三列词表（单词 / 音标 / 词义）变成
    「apple /ˈæpl/ 苹果」：单空格不切分，中文答案原样印在题面第一栏 ——
    听写单成了答案页，音标还会被当释义念出来。
    这里只留「题面 + 释义」两样：序号列、音标列、词性列先剔掉，
    英文取第一个有拉丁字母的单元格，释义取它后面第一个带中文的单元格。
    """
    cleaned = []
    for cell in cells:
        text = str(cell or "").strip()
        if not text:
            continue
        cleaned.append(text)
    if not cleaned:
        return ""
    if len(cleaned) >= 3:
        # 序号列只在多列时才有必要剔（两列词表的第一个格永远是单词，不能按数字剔错）
        filtered = []
        for text in cleaned:
            if RE_INDEX_CELL.match(text):
                continue
            if not RE_CJK.search(text) and (RE_IPA.search(text) or RE_POS_CELL.match(text)):
                continue
            filtered.append(text)
        if filtered:
            cleaned = filtered
    left = cleaned[0]
    gloss = ""
    for text in cleaned[1:]:
        if RE_CJK.search(text) and len(text) <= 30:
            gloss = text
            break
    if gloss and _has_latin(left) and not RE_CJK.search(left):
        return left + " " + gloss
    return left


def _looks_like_header(cells):
    if not cells:
        return False
    hits = 0
    for cell in cells:
        text = str(cell).strip().lower()
        if not text:
            continue
        if any(word in text for word in HEADER_WORDS) and len(text) <= 8:
            hits += 1
    return hits >= max(1, len([cell for cell in cells if str(cell).strip()]))


def _is_word_gloss_pair(text):
    """这一行是不是「英文词 + 中文释义」的正经词条（防止放宽标题判据时把真词条吞掉）。

    "test 测试"、"practice 练习"、"name 姓名"、"look after 照顾" 都是词条：
    左边一个词（最多四个词的短语）、右边一段中文释义，没有数字和句子零件。
    "Unit 3 What does he look like? 单词表" 左边是带数字和问号的句子，不算。
    """
    left, right = _split_line(text)
    if not left or not right:
        return False
    if not _has_latin(left) or RE_CJK.search(left):
        return False                      # 英文半边必须干净（"（Unit 3" 这种是落款碎片）
    if len(right) > 10 or not RE_CJK.search(right):
        return False                      # 释义太长、或根本不是中文：不是词条
    if len(left.split()) > 4 or RE_SENTENCE_MARK.search(left):
        return False                      # 左边是一整句（标题），不是一个词
    return True


def _looks_like_heading(text, position=None):
    """「Unit 4 单词表」「第三单元词汇」这种整行标题不当词条。

    两条判据叠加：
    1. 老规则：含标题词 + 词数不多（"Unit 4 单词表" / "Unit 3 单词听写"）；
    2. 位置规则（position="first"/"last"）：首段、末段只要带“抬头词”（单词表 / 听写 /
       Unit N / 姓名 / 班级 / 日期 / 签名…）且不是「英文 + 中文」两段式，就当标题或落款。
       老师的配套词表正是这个形状：首段 "Unit 3 What does he look like? 单词表"（8 个词）、
       末段 "（Unit 3 单词表 · 七年级英语 · 王老师）"（7 个词）—— 老规则按词数卡死在 4，
       两条都漏过去，印出来就成了第 15、16 个单词。

    「test 测试」「practice 练习」这类右边两个字的真词条必须继续当词条（有测试钉着），
    所以第 2 条一律先过 _is_word_gloss_pair()：两段式真词条直接放行。
    """
    line = str(text or "").strip()
    if not line:
        return False
    if len(line) <= 40 and RE_HEADING.search(line):
        tokens = line.split()
        if len(tokens) <= 2:
            # 「单词表」「Unit 4」「单词听写」这种没配对的，算标题
            pair = (len(tokens) == 2 and _has_latin(tokens[0]) and not RE_CJK.search(tokens[0])
                    and RE_CJK.search(tokens[1]) and not _has_latin(tokens[1]))
            if not pair:
                return True
        elif len(tokens) <= 4:
            return True
    if position in ("first", "last") and len(line) <= 120 and RE_HEADING_WIDE.search(line):
        return not _is_word_gloss_pair(line)
    return False


def _body_blocks(document):
    """按 Word 文档里真实的先后顺序遍历正文块（段落与表格交错），跳过 sectPr 之类。

    为什么不能分开遍历：老师的配套词表是「首段标题 + 两列表格 + 末段落款」。
    先收集所有表格、再把段落 append 到末尾，标题与落款就永远排在词表屁股后面，
    印出来正好成了第 15、16 个单词（评测 P0-2）。只有按文档序走，标题才在第一位，
    "首段/末段才是抬头"这条判据才有意义。
    """
    body = document.element.body
    for child in body.iterchildren():
        tag = child.tag.rsplit("}", 1)[-1]
        if tag == "p":
            yield "paragraph", Paragraph(child, document)
        elif tag == "tbl":
            yield "table", Table(child, document)


def read_word_list(path):
    """从 Word / Excel / CSV / TXT 读单词表，吐成 parse_items 认识的「英文 中文」文本。

    为什么值得做：老师的词表基本不在对话框里——在教材配套的 Excel、备课组共享的 Word、
    或者从某个网站导出的 csv 里。让他们先复制一遍再粘，等于白送一次出错机会。
    """
    require(bool(path), "source-required", "没有给单词表文件", WORD_FILE_HINT)
    full = os.path.abspath(os.fspath(path))
    require(os.path.isfile(full), "file-not-found", "找不到这个文件：" + full, WORD_FILE_HINT)
    ext = os.path.splitext(full)[1].lower()
    lines = []

    if ext in (".xlsx", ".xlsm", ".xltx"):
        from openpyxl import load_workbook
        try:
            book = load_workbook(full, read_only=True, data_only=True)
        except Exception as exc:  # noqa: BLE001 - 三方库异常类型五花八门，统一转中文
            raise EngineError("unsupported-file", "这个表格打不开：" + os.path.basename(full),
                              "确认它是 .xlsx（不是改后缀的 .xls），或者用 Excel 另存一份")
        try:
            for sheet in book.worksheets:
                for row in sheet.iter_rows(values_only=True):
                    cells = []
                    for cell in row or ():
                        text = str(cell).strip() if cell is not None else ""
                        if text:
                            cells.append(text)
                    entry = _row_to_entry(cells)
                    if entry and not _looks_like_header(cells):
                        lines.append(entry)
        finally:
            book.close()
    elif ext in (".docx", ".docm", ".dotx"):
        try:
            document = Document(full)
        except Exception as exc:  # noqa: BLE001
            raise EngineError("unsupported-file", "这个 Word 打不开：" + os.path.basename(full),
                              "用 Word / WPS 另存为 .docx 再来")
        blocks = list(_body_blocks(document))   # 生成器只能走一遍，先落成列表
        # 首段/末段是标题与落款可能出现的位置（中间夹的都是词条），先数清有几段
        paragraphs = [(kind, block) for kind, block in blocks
                      if kind == "paragraph" and (block.text or "").strip()]
        total = len(paragraphs)
        seen = 0
        for kind, block in blocks:
            if kind == "table":
                for row in block.rows:
                    cells = []
                    for cell in row.cells:
                        text = (cell.text or "").strip()
                        if text and text not in cells:
                            cells.append(text)
                    entry = _row_to_entry(cells)
                    if entry and not _looks_like_header(cells):
                        lines.append(entry)
                continue
            text = (block.text or "").strip()
            if not text:
                continue
            position = "first" if seen == 0 else ("last" if seen == total - 1 else None)
            seen += 1
            if _looks_like_header([text]) or _looks_like_heading(text, position):
                continue
            lines.append(text)
    elif ext in (".csv", ".txt"):
        raw = None
        for encoding in ("utf-8-sig", "utf-8", "gbk", "gb18030"):
            try:
                with open(full, "r", encoding=encoding) as handle:
                    raw = handle.read()
                break
            except UnicodeDecodeError:
                continue
        require(raw is not None, "unsupported-file", "这个文本文件的编码认不出来",
                "用 Excel / 记事本另存为 UTF-8 或者 GBK 再来")
        if ext == ".csv":
            import csv
            # io.StringIO 才认引号里的换行；raw.splitlines() 会把一条多行单元格劈成两条
            for row in csv.reader(io.StringIO(raw, newline="")):
                entry = _row_to_entry(row)
                if entry and not _looks_like_header([str(cell).strip() for cell in row if str(cell).strip()]):
                    lines.append(entry)
        else:
            lines.extend(text for text in raw.splitlines() if text.strip())
    else:
        raise EngineError("unsupported-source", "这个格式我读不了（" + (ext or "没有扩展名") + "）",
                          "存成 .docx / .xlsx / .csv / .txt 再来")

    text = "\n".join(lines).strip()
    require(bool(text), "no-items", "这个文件里没读出文字：" + os.path.basename(full),
            "确认里面有单词表（Word 表格、Excel 两列、或者一行一条的 txt）")
    return text


def build_dictation(payload):
    """payload = {text 或 source, title?, mode?, columns?, count?, answer?, blankLength?,
                  lineSpacing?, page?, baseTemplate?, outputPath?, overwrite?, values?}

    source：给一个 .docx / .xlsx / .csv / .txt，代替 text（text 为空时才用）。

    返回 {outputPath, count, mode, columns, warnings, page, bytes}
    """
    payload = payload or {}
    text = payload.get("text")
    source = payload.get("source")
    if (not str(text or "").strip()) and source:
        text = read_word_list(source)
    require(bool(str(text or "").strip()), "text-required", "没有收到听写内容",
            TEXT_HINT + "；也可以点「浏览…」直接选一个词表文件")

    mode = _norm_mode(payload.get("mode"))
    columns = _to_int(payload.get("columns", DEFAULT_COLUMNS), "bad-columns",
                      "分栏数只能是 1—" + str(MAX_COLUMNS) + "：" + str(payload.get("columns")),
                      "听写单通常 2 栏", 1, MAX_COLUMNS)
    blank_length = _to_int(payload.get("blankLength", DEFAULT_BLANK_LENGTH), "bad-blank-length",
                           "横线长度不合理：" + str(payload.get("blankLength")),
                           "给 1—200 个下划线，默认 14", 1, 200)
    line_spacing = _to_float(payload.get("lineSpacing", DEFAULT_LINE_SPACING), "bad-line-spacing",
                             "行距不合理：" + str(payload.get("lineSpacing")),
                             "给 0.5—5 倍，默认 2.0（写完一行留得下手写高度）", 0.5, 5.0)
    title = str(payload.get("title") or DEFAULT_TITLE).strip() or DEFAULT_TITLE
    with_answer = bool(payload.get("answer", False))

    parsed = parse_items(text, mode)
    items = parsed["items"]
    warnings = list(parsed["warnings"])
    require(bool(items), "no-items", "这段文字里一条听写都没认出来",
            "检查是不是整段连在一起了；" + TEXT_HINT)

    count = payload.get("count")
    if count is not None:
        wanted = _to_int(count, "bad-count", "条目数不合理：" + str(count),
                         "给 1—" + str(MAX_ITEMS), 1, MAX_ITEMS)
        if wanted > len(items):
            warnings.append({
                "code": "not-enough-items",
                "message": "只拆出 " + str(len(items)) + " 条，比要求的 " + str(wanted) + " 条少",
            })
        items = items[:wanted]

    # 容量估算：一页 A4 单栏大约 24 行，双栏 48 条。
    # 40 个词 2 栏会印成 2 页、尾页可能只剩三四条 —— 不提示的话老师要不发现少了一页，
    # 要不临上课才发现最后一行孤零零挂着。
    per_column = 24
    capacity = columns * per_column
    pages = max(1, (len(items) + capacity - 1) // capacity)
    if pages > 1:
        tail = len(items) - (pages - 1) * capacity
        warnings.append({
            "code": "spill-pages",
            "message": "共 " + str(len(items)) + " 条，按 " + str(columns) + " 栏大约印 " + str(pages)
                       + " 页（每页约 " + str(capacity) + " 条），最后一页只有 " + str(tail) + " 条；"
                       + "想更紧凑可以改 1 栏或把 count 调小，想让尾页不空就补几条",
        })

    blocks = [
        {"type": "paragraph", "align": "center",
         "runs": [{"text": title, "bold": True, "sizePt": 16}]},
        {"type": "paragraph", "align": "center",
         "runs": [{"text": INFO_LINE, "sizePt": 10.5}]},
    ]
    # 标题和说明行要跨整页（双栏下不切断，否则"班级/姓名"被挤进左栏一半），
    # 所以先在这两行之后插一个连续分节符：前面单栏，后面按 columns 栏排。
    if columns > 1:
        blocks.append({"type": "columns", "count": 1})
    for item in items:
        blocks.append(_item_block(item, blank_length))

    if with_answer:
        # 参考答案另起一页。分栏时再切一节：让「参考答案」四个字跨整页居中，
        # 答案条目仍然按 columns 栏排（和正文同一套版式，老师看着是一份东西）。
        if columns > 1:
            blocks.append({"type": "columns", "count": columns})
        blocks.append({"type": "paragraph", "pageBreakBefore": True, "align": "center",
                       "runs": [{"text": ANSWER_HEADING, "bold": True, "sizePt": 14}]})
        if columns > 1:
            blocks.append({"type": "columns", "count": 1})
        for item in items:
            blocks.append({"type": "paragraph", "runs": [{"text": _answer_line(item, mode)}]})

    page = payload.get("page") or {}
    if isinstance(page, str):
        page = {"size": page}  # 「A4」这种老板式写法：别的命令都认，听写单不该单独报错
    elif not isinstance(page, dict):
        raise EngineError("bad-page", "纸型要写成 A4 / 8K，或者 {size, orientation} 对象")
    page = dict(page)
    page.setdefault("size", "A4")
    page.setdefault("orientation", "portrait")
    page["columns"] = columns      # 栏数以 columns 为准，免得两处不一致印出来对不上

    output = payload.get("outputPath") or default_output("听写", title, ".docx")
    spec = {
        "outputPath": output,
        "page": page,
        "blocks": blocks,
        "values": dict(payload.get("values") or {}),
        "properties": {"title": title},
        "overwrite": bool(payload.get("overwrite", False)),
    }
    if payload.get("baseTemplate"):
        spec["baseTemplate"] = payload["baseTemplate"]
    built = build_docx(spec)
    _tighten_items(built["outputPath"], line_spacing)

    result_path = built["outputPath"]
    warnings.extend(list(built.get("warnings") or []))
    return {
        "outputPath": result_path,
        "count": len(items),
        "mode": mode,
        "columns": columns,
        "warnings": warnings,
        "page": built.get("page"),
        "bytes": os.path.getsize(result_path),
    }

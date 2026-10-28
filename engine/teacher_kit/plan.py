"""教案：把老师粘进来的一段备课文字，排成能直接交检的 Word 教案。

为什么这样排（电子教案绝大多数是「迎检」用的，不是给自己看的）：
- 顶部固定一块「课题｜课型｜授课教师｜班级」信息表：检查组第一眼就看这四个格子，
  而且学校模板里这四项永远在同一位置，位置一漂就被当成「换了格式」退回。
- 正文一块字段一张表（字段名一行、内容一行）：有框线才像学校发的表，老师还能在
  空行上手写补充；纯段落排版在迎检时反而被说「不像教案表」。
- 「教学过程」必须是 4 列表（环节｜教师活动｜学生活动｜时间）：这是检查组唯一会
  逐格看的地方 —— 只写成段落，他们没法确认「学生活动」「时间」到底写了没有。
- 「板书设计」永远是 1×1 带框格子：板书是画出来的，老师要在格子里手写或贴图，
  做成正文段落等于没给地方。
- A4 竖版：电子教案一律 A4 归档；8 开横向是卷子的排法，教案套卷子版式会被要求返工。

解析上宁可少猜：认不出的行不丢、不硬塞，一律进「教学过程」并给一条中文提示 ——
老师粘进来的字一个都不能少，少了就得重打一遍。
"""

import os
import re

import docx
from docx.enum.table import WD_ALIGN_VERTICAL, WD_ROW_HEIGHT_RULE
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor, Twips

from .docx_build import build_docx, column_width_twips
from .errors import EngineError, require
from .paths import default_output
from .placeholders import fill_document

# 教案自己的页脚。docx_build 的默认页脚是卷子版（「出卷老师」），
# 教案上印「出卷老师」会被检查组当成拿卷子模板糊的，所以这里自己写一行。
PLAN_FOOTER = "第 {{page}} 页 共 {{pages}} 页　　授课教师：__________"

CN_NUM = "一二三四五六七八九十"

# 字段表：key 给机器认，label 给老师看，aliases 是各种教案模板里真实出现过的写法。
# 顺序就是成品从上到下的顺序 —— 迎检的检查顺序基本也是这个顺序
# （课题 → 目标 → 重点 → 难点 → 过程 → 板书 → 作业 → 反思）。
FIELDS = (
    {"key": "title", "label": "课题", "aliases": ("课题", "教学内容", "授课内容", "教学课题", "课题名称")},
    {"key": "lessonType", "label": "课型", "aliases": ("课型", "课的类型", "课程类型", "授课类型")},
    {"key": "periods", "label": "课时", "aliases": ("课时", "授课课时", "课时安排", "教学课时")},
    {"key": "className", "label": "班级", "aliases": ("班级", "授课班级", "教学班级", "任课班级")},
    # 只认「授课教师」这一族全称：老师写教学过程时常写「教师：…」，
    # 把光秃秃的「教师」也当字段名，教学过程会被从中间劈开、后半段全跑进信息表。
    {"key": "teacher", "label": "授课教师",
     "aliases": ("授课教师", "任课教师", "执教教师", "主讲教师", "上课教师", "教师姓名")},
    {"key": "studentAnalysis", "label": "学情分析",
     "aliases": ("学情分析", "学情", "学生分析", "学生情况分析", "学情研判")},
    {"key": "objectives", "label": "教学目标",
     "aliases": ("教学目标", "学习目标", "教学目的", "教学要求", "教学目标与要求", "教学目标及要求")},
    {"key": "keyPoints", "label": "教学重点",
     "aliases": ("教学重点", "重点", "教学重难点", "教学重点难点", "重难点")},
    {"key": "difficultPoints", "label": "教学难点", "aliases": ("教学难点", "难点")},
    {"key": "preparation", "label": "教学准备",
     "aliases": ("教学准备", "教具", "教具准备", "教学用具", "准备", "课前准备")},
    {"key": "procedure", "label": "教学过程",
     "aliases": ("教学过程", "教学流程", "教学过程设计", "课堂过程", "教学活动")},
    {"key": "boardDesign", "label": "板书设计", "aliases": ("板书设计", "板书", "板书内容")},
    {"key": "homework", "label": "作业布置",
     "aliases": ("作业布置", "作业", "课后作业", "布置作业", "作业设计")},
    {"key": "reflection", "label": "教学反思",
     "aliases": ("教学反思", "课后反思", "反思", "教学反思与改进")},
)

FIELD_BY_KEY = {spec["key"]: spec for spec in FIELDS}
# 顶部信息表里的四项：它们只在上表出现，不在正文再排一遍（同一份教案印两遍课题最刺眼）
# 这几项进顶部信息表，正文里不再重复印一遍。
# 课时也放进来：它只有一个短值，单独排成「字段名 / 值」两块会被看成「值掉到下一行」。
HEADER_KEYS = ("title", "lessonType", "teacher", "className", "periods")
PROCEDURE_KEY = "procedure"
BOARD_KEY = "boardDesign"

# 学生活动/时间留白时跟在教学过程表后面的一行小字。
# 空白的格子必须是「明确留给你填」，不能看着像生成失败 —— 但也不能瞎编学生的反应。
NOTE_BLANK = ("注：「学生活动」「时间」里空着的格子是留给你按本班实际写的 —— "
              "自动编的活动容易和学生的真实反应对不上。")

# 别名 → key。别名统一去掉空格后比对，所以「教学 目标」也能认。
ALIAS_INDEX = {}
for _spec in FIELDS:
    for _alias in (_spec["label"],) + tuple(_spec["aliases"]):
        ALIAS_INDEX[re.sub(r"[\s\u3000]+", "", _alias)] = _spec["key"]

MAX_UNKNOWN_WARNINGS = 20  # 一行一条提示；整篇没标题时别把返回刷成 50 条

# ── 标记行 ────────────────────────────────────────────────────────────────
RE_BRACKET = re.compile(r"^\s*[【\[]\s*([^】\]]{1,24})\s*[】\]]\s*(.*)$")
RE_MD_HEAD = re.compile(r"^\s*#{1,6}\s*(.+?)\s*#*\s*$")
RE_CN_PREFIX = re.compile(r"^\s*([" + CN_NUM + r"]{1,3})\s*[、.．]\s*(.*)$")
RE_COLON = re.compile(r"^\s*([^:：]{1,24}?)\s*[:：]\s*(.*)$")
RE_PAREN_SUFFIX = re.compile(r"[（(][^）)]*[）)]$")

# ── 教学过程环节 ─────────────────────────────────────────────────────────
# 学校电子教案里环节标题的常见写法：一、导入 / （一）导入 / 环节一：导入 / Step 1: 导入 / 导入：
SECTION_WORDS = (
    "导入", "导入新课", "情境导入", "创设情境", "热身", "复习", "复习导入", "预习",
    "新授", "讲授", "新授知识", "精讲", "点拨", "讲解",
    # 备课稿里最常见的几个环节名，之前一个都不在表里 —— 整段被并进上一环节
    "呈现", "新知", "新知呈现", "新课呈现", "新课学习", "学习新课", "拓展", "延伸",
    "拓展延伸", "输出", "语言输出", "学以致用", "迁移", "应用",
    "操练", "练习", "巩固", "巩固练习", "当堂检测", "检测", "反馈", "讲评",
    "讨论", "探究", "合作", "展示", "活动", "任务", "游戏",
    "小结", "总结", "课堂小结", "归纳",
    "作业", "布置作业", "作业布置",
    "板书", "板书设计",
    "warm-up", "warming up", "lead-in", "presentation", "practice", "production",
    "summary", "homework", "revision",
)
SECTION_WORD_SET = {word.lower() for word in SECTION_WORDS}

RE_SEC_PAREN = re.compile(r"^[（(]\s*(?P<num>[" + CN_NUM + r"]{1,3}|\d{1,2})\s*[)）]\s*(?P<rest>.*)$")
RE_SEC_CN = re.compile(r"^(?P<num>[" + CN_NUM + r"]{1,3})\s*[、.．]\s*(?P<rest>.*)$")
RE_SEC_HUANJIE = re.compile(
    r"^第?\s*(?P<num>[" + CN_NUM + r"\d]{1,3})?\s*(?:个)?环节\s*"
    r"(?P<num2>[" + CN_NUM + r"\d]{1,3})?\s*[:：、.．]?\s*(?P<rest>.*)$")
RE_SEC_STEP = re.compile(r"^(?i:step)\s*(?P<num>\d{1,2})\s*[:：、.．)）]?\s*(?P<rest>.*)$")
# 阿拉伯编号（1. 导入）只在标题里带环节关键词时才当环节：
# 「1. 学生读课文」是过程里的一条细项，不是环节标题，切错了表格会碎成一格一行。
RE_SEC_ARABIC = re.compile(r"^(?P<num>\d{1,2})\s*[、.．)）]\s*(?P<rest>.*)$")

RE_TIME_PAREN = re.compile(r"[（(]\s*([^（）()]{1,16})\s*[)）]\s*$")
RE_TIME_HINT = re.compile(r"\d+\s*(?:分钟|分|min)", re.IGNORECASE)
# 「操练 12 分钟」这种不带括号的写法：空格后面直接跟时长
RE_TIME_BARE = re.compile(r"^(?P<name>.+?)[\s\u3000]+(?P<time>\d{1,3}\s*(?:分钟|分|min|mins))$",
                          re.IGNORECASE)
# 裸环节名（没编号、没冒号）最多这么长：「导入」「新课导入」「词汇操练」都在范围里，
# 再长就不是环节名而是正文句子了 —— 切错一刀，4 列表就碎成一格一行。
MAX_BARE_SECTION_CHARS = 12
RE_SENTENCE_END = re.compile(r"[。！？!?；;]$")

# 「教师：…」「【学生活动】…」「时间：5 分钟」——三类活动标签，决定内容落到哪一格
ACTIVITY_PATTERNS = (
    (re.compile(r"^(?:【\s*(?:教师活动|教师|老师|师)\s*】|(?:教师活动|教师|老师|师)\s*[:：])\s*(.*)$"), "teacher"),
    (re.compile(r"^(?:【\s*(?:学生活动|学生|学|生)\s*】|(?:学生活动|学生|学|生)\s*[:：])\s*(.*)$"), "student"),
    (re.compile(r"^(?:【\s*(?:时间|用时|时长)\s*】|(?:时间|用时|时长)\s*[:：])\s*(.*)$"), "time"),
)
BUCKET_KEY = {"teacher": "teacher", "student": "student", "time": "times"}


def _normalize_head(text):
    """把标记行的小标题洗成能查别名的样子：去空白、去包裹符号、去句尾标点。"""
    value = re.sub(r"[\s\u3000]+", "", str(text or ""))
    value = value.strip("【】[]「」〈〉<>《》")
    return value.strip(":：、。.．,，;；")


def _head_key(text):
    """小标题 → 字段 key；认不出返回 None（认不出就不猜）。"""
    value = _normalize_head(text)
    if not value:
        return None
    direct = ALIAS_INDEX.get(value)
    if direct:
        return direct
    # 「教学目标（三维）」「教学重点(1)」这种带括号补充的，去掉括号再认一次
    return ALIAS_INDEX.get(RE_PAREN_SUFFIX.sub("", value).strip())


def _clean_inline(text):
    """标记行冒号后面的内容：「教学目标：1. …」里的「1. …」。"""
    value = str(text or "").strip(" \u3000")
    value = re.sub(r"^[:：]\s*", "", value)
    return value.strip()


def _split_colon(text):
    match = RE_COLON.match(str(text or ""))
    if not match:
        return str(text or "").strip(), ""
    return match.group(1).strip(), match.group(2).strip()


def _marker(line):
    """这一行是字段标记吗？返回 (key, 行内正文) 或 None。

    认四种写法：课题：… / 一、教学目标 / 【教学目标】 / [教学目标] / ## 教学目标。
    只认「整行就是字段名（可带冒号）」—— 正文里出现「教学目标」四个字不会被切开。
    """
    text = str(line or "").strip().strip("\u3000")
    if not text:
        return None
    bracket = RE_BRACKET.match(text)
    if bracket:
        key = _head_key(bracket.group(1))
        if key:
            return key, _clean_inline(bracket.group(2))
    markdown = RE_MD_HEAD.match(text)
    if markdown and markdown.group(1):
        key = _head_key(markdown.group(1))
        if key:
            return key, ""
    cn_prefix = RE_CN_PREFIX.match(text)
    if cn_prefix:
        head, inline = _split_colon(cn_prefix.group(2))
        key = _head_key(head)
        if key:
            return key, _clean_inline(inline)
    head, inline = _split_colon(text)
    key = _head_key(head)
    if key:
        return key, _clean_inline(inline)
    return None


def _split_time(text):
    """「导入（5 分钟）」→ ("导入", "5 分钟")；「操练 12 分钟」→ ("操练", "12 分钟")。

    括号里不是时间的（比如「（一）」「（拓展）」）原样留着，交给后面的环节名判断。
    """
    value = str(text or "").strip()
    match = RE_TIME_PAREN.search(value)
    if match:
        inner = match.group(1).strip()
        if not RE_TIME_HINT.search(inner):
            return value, ""
        return value[:match.start()].strip(), inner
    bare = RE_TIME_BARE.match(value)
    if bare:
        return bare.group("name").strip(), bare.group("time").strip()
    return value, ""


def _is_section_name(text):
    """裸环节名判断：「导入」「新课导入」「词汇操练」「Practice」算，「复习单词」不算。

    老师最自然的写法是「导入（5 分钟）」「新授(10min)」，前面既没有「一、」也没有冒号。
    判据只有两条：
    - 去掉时长后整串就是一个环节关键词（导入/新授/操练/巩固/小结/作业/板书…）；
    - 或者结尾就是环节关键词、前面只多出两三个修饰字（新课导入 / 词汇操练 / 课堂小结）。
    按「结尾是环节词」判而不是「开头是环节词」：正文里随手写的「复习单词」「操练句型」
    结尾是「单词」「句型」，不会被切成一格一行的碎表。
    """
    value = _normalize_head(text).lower()
    if not value:
        return False
    if value in SECTION_WORD_SET:
        return True
    if len(value) > MAX_BARE_SECTION_CHARS:
        return False
    for word in SECTION_WORD_SET:
        # 英文只认整串相等：「practice speaking」是正文，不是小标题
        if ord(word[0]) <= 0x2E80:
            continue
        if value.endswith(word) and 1 <= len(value) - len(word) <= 4:
            return True
    return False


def _procedure_start(line):
    """这一行是不是「教学过程」的开头（环节标题里带环节关键词）？

    很多老师干脆不写「教学过程」四个字，直接从「一、导入」往下写。
    这时候如果照搬「归到当前字段」的规则，整段过程会塞进教学目标里 ——
    所以只要有环节标题、标题里又带环节关键词（导入/新授/操练…），就认作教学过程开始。
    只认关键词，是为了不把「一、知识目标」「二、能力目标」这种目标分层当成环节。
    """
    head = _section_head(line)
    if not head:
        return False
    title = _normalize_head(head[0]).lower()
    if not any(word in title for word in SECTION_WORD_SET):
        return False
    if head[3] == "arabic":
        # 「1. 复习单词」是过程里的一条细项，不当成过程开始；
        # 只有「1. 导入」这种标题就等于环节关键词的才算
        return title in SECTION_WORD_SET
    return True


def _starts_with_section_word(text):
    lowered = _normalize_head(text).lower()
    return any(lowered.startswith(word) for word in SECTION_WORD_SET)


def _section_head(line):
    """这一行是环节标题吗？返回 (标题, 时间, 行内正文) 或 None。

    认四种写法：
    - 带编号：「一、导入」「（一）导入」「环节一：导入」「Step 2: Practice」「1. 导入」；
    - 环节关键词 + 冒号：「导入：」；
    - 裸环节名（P1-7）：「导入」「导入（5 分钟）」「新授(10min)」「操练 12 分钟」「小结」——
      教研组最常这么写，之前只认前两种，整段过程被挤进一格；
    - 上面任意一种后面跟时长：括号里的、或者空格直接跟「N 分钟 / Nmin」的，都进「时间」列。

    过程里随手写的一句话（"学生先读课文"、"复习单词三遍"）不会被当成环节标题，
    否则 4 列表会被切成一堆一格一行的碎表。
    """
    text = str(line or "").strip().strip("\u3000")
    if not text or len(text) > 80:
        return None
    title = None
    inline = ""
    time_text = ""
    kind = "label"
    for pattern in (RE_SEC_PAREN, RE_SEC_HUANJIE, RE_SEC_STEP, RE_SEC_CN):
        match = pattern.match(text)
        if not match:
            continue
        title, inline = _split_colon(match.group("rest") or "")
        if not title:
            return None
        break
    if title is None:
        arabic = RE_SEC_ARABIC.match(text)
        if arabic:
            head, tail = _split_colon(arabic.group("rest") or "")
            if not _starts_with_section_word(head):
                return None
            kind = "arabic"
            title, inline = head, tail
        else:
            head, tail = _split_colon(text)
            name, bare_time = _split_time(head)
            if not _is_section_name(name):
                return None
            title, inline, time_text = name, tail, bare_time
    title, more_time = _split_time(title.strip())
    if more_time:
        time_text = time_text or more_time
    # 句末标点必须在 strip 之前判：标题末尾的「。」说明这是一句话，不是环节名
    if not title or len(title) > 20 or RE_SENTENCE_END.search(title):
        return None
    title = title.strip(" 　:：,，、。.．")
    if not title:
        return None
    return title, time_text, _clean_inline(inline), kind


def _route_line(section, line):
    """把一行内容放进教师活动/学生活动/时间那一格；没贴标签就跟着上一行的那一格。"""
    for pattern, bucket in ACTIVITY_PATTERNS:
        match = pattern.match(line)
        if not match:
            continue
        section["bucket"] = bucket
        value = _clean_inline(match.group(1))
        if value:
            section[BUCKET_KEY[bucket]].append(value)
        return bucket
    bucket = section.get("bucket") or "teacher"
    if bucket == "time":
        # 上一行是「时间：5 分钟」时，没贴标签的后续行会跟着进「时间」格 ——
        # 环节标题就是这么被塞进 1100 twips 窄列、印成竖排的。没标签的一律回落到教师活动。
        bucket = "teacher"
    section[BUCKET_KEY[bucket]].append(line)
    return bucket


def _split_procedure(lines):
    """教学过程 → 环节列表。每个环节记原文行、教师活动、学生活动、时间。

    原文行（lines）保留给 flow 版式和面板预览用；三份活动清单给 4 列表用。
    环节目录里没写「教师：」「学生：」的，整段归教师活动 —— 那是老师自己心里的话，
    空着学生活动那一格让他手写，比瞎猜一句「学生听讲」强；
    但空着的格子必须在成品里标出来（表头写「老师自填」+ 表下一行说明），
    否则老师第一眼会以为「这一列没生成」。
    """
    sections = []
    current = None
    for line in lines:
        head = _section_head(line)
        if head:
            title, time_text, inline = head[0], head[1], head[2]
            current = {"title": title, "times": [], "lines": [], "teacher": [], "student": [], "bucket": None}
            if time_text:
                current["times"].append(time_text)
            sections.append(current)
            if inline:
                current["lines"].append(inline)
                current["teacher"].append(inline)
                current["bucket"] = "teacher"
            continue
        if current is None:
            # 第一个环节标题之前的内容：给一行空环节，老师手写个标题就成表
            current = {"title": "", "times": [], "lines": [], "teacher": [], "student": [], "bucket": None}
            sections.append(current)
        bucket = _route_line(current, line)
        # 「时间：3 分钟」只进时间格：留在原文行里，整段版（flow）会把它当正文再印一遍，
        # 时间就出现两次（「（导入）　3 分钟」+ 正文「时间：3 分钟」）。
        if bucket != "time":
            current["lines"].append(line)
    return sections


def _read_plan(text):
    """解析主流程：标记行切字段，剩下的行归当前字段；谁都不认的行进教学过程。"""
    require(isinstance(text, str), "text-not-string", "教案要是一段文字（把备课内容粘贴进来）", "把课题、教学目标、教学过程粘成一段纯文字，表格形式的教案也能认")
    require(bool((text or "").strip()), "text-required", "没有收到教案内容", "把课题、教学目标、教学过程粘进来就行")

    lines = text.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    collected = {}
    unknown = []
    duplicates = []
    inferred = None  # 「教学过程」没写标题、但被环节行带出来的行号
    current = None

    for line_number, raw in enumerate(lines, start=1):
        line = raw.strip().strip("\u3000")
        if not line:
            # 空行是老师排版的空行，不是「认不出」—— 为它刷提示只会让人以为出错了
            continue
        marker = _marker(line)
        if marker:
            key, inline = marker
            if key in collected:
                duplicates.append((line_number, key))
            else:
                collected[key] = []
            current = key
            if inline:
                collected[key].append(inline)
            continue
        if current is None:
            unknown.append((line_number, line))
            continue
        if current != PROCEDURE_KEY and _procedure_start(line):
            if PROCEDURE_KEY not in collected:
                collected[PROCEDURE_KEY] = []
                inferred = line_number
            current = PROCEDURE_KEY
        collected[current].append(line)

    warnings = []
    if inferred is not None:
        warnings.append({
            "line": inferred,
            "code": "procedure-inferred",
            "message": "第 " + str(inferred) + " 行起像是「教学过程」的环节，但这块没写标题，已自动归到教学过程",
        })
    for line_number, key in duplicates:
        warnings.append({
            "line": line_number,
            "code": "duplicate-field",
            "message": "第 " + str(line_number) + " 行又出现了一次「" + FIELD_BY_KEY[key]["label"]
                       + "」，已合并到前面那一块（不重复印两张表）",
        })

    # 备课稿第一行常常就是课题（「Unit 5 现在完成时」），没写「课题：」；
    # 老逻辑把它塞进教学过程 → 大标题印成「教案」、课题格空白。这里先捞出来。
    if "title" not in collected:
        for position, (line_number, line) in enumerate(unknown):
            text = str(line or "").strip()
            if not text or len(text) > 40:
                continue
            if RE_SENTENCE_END.search(text) or "：" in text or ":" in text:
                continue
            if _section_head(text):
                continue
            collected["title"] = [text]
            unknown.pop(position)
            warnings.append({
                "line": line_number,
                "code": "title-inferred",
                "message": "第 " + str(line_number) + " 行看着像课题，已当课题处理（想改就写成「课题：…」）",
            })
            break

    # 认不出归属的行一律塞进教学过程：宁可结构难看，也不能把老师写的字丢掉
    if unknown:
        collected.setdefault(PROCEDURE_KEY, [])
        collected[PROCEDURE_KEY] = [line for _, line in unknown] + collected[PROCEDURE_KEY]
        for line_number, _ in unknown[:MAX_UNKNOWN_WARNINGS]:
            warnings.append({
                "line": line_number,
                "code": "unknown-line",
                "message": "第 " + str(line_number) + " 行没认出属于哪一块，已放进教学过程",
            })
        if len(unknown) > MAX_UNKNOWN_WARNINGS:
            warnings.append({
                "code": "unknown-line-more",
                "message": "还有 " + str(len(unknown) - MAX_UNKNOWN_WARNINGS)
                           + " 行也没认出属于哪一块，同样放进了教学过程",
            })

    fields = {}
    for spec in FIELDS:
        key = spec["key"]
        if key in collected:
            fields[key] = "\n".join(collected[key]).strip()

    warnings.sort(key=lambda item: (item.get("line") is None, item.get("line") or 0))
    return {
        "fields": fields,
        "sections": _split_procedure(collected.get(PROCEDURE_KEY) or []),
        "warnings": warnings,
    }


def parse_plan(payload):
    """把老师粘的一段教案拆成字段与教学过程环节。

    payload 可以是字符串，也可以是 {text: ...}（命令表按后者调用）——
    之前只收字符串，plan:parse 经 CLI 走一律是 dict，命令 100% 报错。

    返回 {fields:{key: 原文}, sections:[{title, lines}], warnings:[{code, message, line?}]}。
    fields 里的值按行原样保留（目标的分层编号 1./2./3. 以及「知识/能力/情感」都不动）。
    """
    if isinstance(payload, dict):
        text = payload.get("text")
    else:
        text = payload
    parsed = _read_plan(text)
    return {
        "fields": parsed["fields"],
        "sections": [{"title": section["title"], "lines": list(section["lines"])}
                     for section in parsed["sections"]],
        "warnings": parsed["warnings"],
    }


def plan_fields(_payload=None):
    """面板将来用：有哪些字段、叫什么、还认哪些别名。"""
    return [{"key": spec["key"], "label": spec["label"], "aliases": list(spec["aliases"])}
            for spec in FIELDS]


# ── 排版 ──────────────────────────────────────────────────────────────────
def _pick_title(payload, fields):
    for candidate in (payload.get("title"), fields.get("title")):
        text = str(candidate or "").strip()
        if text:
            return text
    return "教案"


def _heading_text(title, subject):
    """首行大标题：「英语《What does he look like?》教案」。已经带「教案」二字的不再套书名号。

    学科只在课题里没有时前缀：「七年级英语 Module 5 Unit 1」再前缀一次，
    大标题就印成「英语《七年级英语 Module 5 Unit 1》教案」，学科出现两遍。
    """
    text = str(title or "").strip() or "教案"
    heading = text if (text.endswith("教案") or text.endswith("教学设计")) else "《" + text + "》教案"
    subject = str(subject or "").strip()
    if not subject or subject.lower() in text.lower():
        return heading
    return subject + " " + heading


def _one_line(text):
    return re.sub(r"\s*\n\s*", " ", str(text or "")).strip()


def _lines(text):
    return [line.strip() for line in str(text or "").split("\n") if line.strip()]


def _clean_times(times):
    """时间列：去重、只留前两条、每条截到 12 字。

    「时间」列宽 1100 twips，塞三条就折成竖排、把表格撑高一大截；
    重复的「5 分钟」还会让人以为环节时长统计有错。
    """
    seen = []
    for raw in times or []:
        text = str(raw or "").strip().strip("，,、;；")
        if text and text not in seen:
            seen.append(text)
    out = []
    for text in seen[:2]:
        out.append(text if len(text) <= 12 else text[:12])
    return out


# 教学过程一个环节装多少字就换行（「（续）」行）。
# 教学过程中列宽约 5.1cm、10.5pt 一行约 13 个汉字，一页能排 40 多行；
# 超过这个数就拆行，保证「一行的内容不会比一页还高」——
# 行高一旦超过一页，Word/LibreOffice 只能从中间切开，续页第一列（环节）就是空白、
# 表头也不会重现（0.18 审查 L-01 的现场）。
CHARS_PER_ROW = 240


def _chunk_lines(lines, limit=CHARS_PER_ROW):
    """把一节的正文按字数拆成几块，每块做成一行表格。"""
    chunks, current, size = [], [], 0
    for raw in lines:
        line = str(raw or "")
        while len(line) > limit:
            if current:
                chunks.append(current)
                current, size = [], 0
            chunks.append([line[:limit]])
            line = line[limit:]
        if current and size + len(line) > limit:
            chunks.append(current)
            current, size = [], 0
        if line:
            current.append(line)
            size += len(line)
    if current:
        chunks.append(current)
    return chunks or [[""]]


def _procedure_rows(sections):
    """教学过程 → 4 列表的行；顺带给出「要不要合并第一列」和三列各填了几格。

    认不出环节时只有一行：整段话连第一列都放不下（环节列窄），
    所以把这一行的后三列并进来 —— 内容还是从第一列开始，但不会被挤成竖排。

    学生活动、时间常常只有老师自己知道，空着是对的（自动编的活动和班上的实际不符）；
    但「整列空着」必须一眼看出来是留白，所以整列都空时表头改写成「学生活动（老师自填）」，
    并把这列的 filled/blank 数一起带回调用方 —— 面板要能说清「不是没生成」。
    """
    merged = (not sections) or all(not section["title"] for section in sections)
    if merged:
        text = "\n".join(line for section in sections for line in section["lines"])
        rows = [[text, "", "", ""]]
        # 合并成一行时后三列本来就是留白，统计要如实报（以前全 0，面板与回执看着像没生成），
        # 也才会跟着印 NOTE_BLANK 那句「空着的格子是留给你写的」。
        stats = {"rows": 1, "teacher": (1, 0), "student": (0, 1), "time": (0, 1)}
    else:
        rows = []
        counts = {"teacher": [0, 0], "student": [0, 0], "time": [0, 0]}
        for section in sections:
            student = "\n".join(section["student"])
            clock = "，".join(_clean_times(section["times"]))
            for index, chunk in enumerate(_chunk_lines(section["teacher"])):
                # 拆出来的续行第一列写「新授（续）」：续页第一列空着，看着像整行被撕断了
                title = section["title"]
                if index and title:
                    title += "（续）"
                rows.append([
                    title if index == 0 or title else "",
                    "\n".join(chunk),
                    student if index == 0 else "",
                    clock if index == 0 else "",
                ])
            # 统计按「环节」算，不按拆出来的行算：续行本来就不该重复报留白
            row = [section["title"], "\n".join(section["teacher"]), student, clock]
            for index, column in ((1, "teacher"), (2, "student"), (3, "time")):
                # [0] 记得是「填了」，[1] 是「留白」
                counts[column][0 if row[index].strip() else 1] += 1
        stats = {"rows": len(rows),
                 "teacher": tuple(counts["teacher"]),
                 "student": tuple(counts["student"]),
                 "time": tuple(counts["time"])}
    # 整列都没写才在表头点明「老师自填」；只空个别格子不点（那是正常留白）
    head = "学生活动（老师自填）" if (stats["rows"] and stats["student"][0] == 0) else "学生活动"
    return [["环节", "教师活动", head, "时间"]] + rows, merged, stats


def _form_blocks(blocks, roles, fields, sections, header, title, subject):
    """表格版：一块字段一张表，教学过程 4 列，板书永远是带框格子。"""
    blocks.append({
        "type": "paragraph",
        "align": "center",
        "runs": [{"text": _heading_text(title, subject), "bold": True, "sizePt": 16}],
    })
    blocks.append({
        "type": "table",
        "tableStyle": "Table Grid",
        "rows": [
            # 课题独占一行（课题最长，半宽格子怎么缩字号都放不下），值和框线在 _polish 里跨三列
            ["课题", header["title"], "", ""],
            ["课型", header["lessonType"], "课时", header["periods"]],
            ["授课教师", header["teacher"], "班级", header["className"]],
        ],
    })
    roles.append({"role": "header"})

    stats = {"rows": 0, "teacher": (0, 0), "student": (0, 0), "time": (0, 0)}
    field_count = 0
    for spec in FIELDS:
        key = spec["key"]
        if key in HEADER_KEYS:
            continue
        if key == BOARD_KEY:
            # 板书框无条件出：老师没写板书，也总要在格子里画一遍（迎检要看见「板书设计」这块）
            blocks.append({"type": "table", "tableStyle": "Table Grid", "rows": [["板书设计"]],
                           "noRowSplit": True})
            roles.append({"role": "board", "lines": _lines(fields.get(BOARD_KEY))})
            continue
        text = str(fields.get(key) or "").strip()
        if not text:
            continue
        field_count += 1
        if key == PROCEDURE_KEY:
            rows, merged, stats = _procedure_rows(sections)
            blocks.append({"type": "table", "tableStyle": "Table Grid", "rows": rows,
                           "repeatHeader": True, "noRowSplit": True})
            roles.append({"role": "procedure", "merged": merged})
            if stats["student"][1]:
                # 有格子留白就跟一行小字说明：老师打印出来第一眼要能看懂「这是给我填的」
                blocks.append({"type": "paragraph", "runs": [{"text": NOTE_BLANK, "sizePt": 9}]})
            continue
        blocks.append({"type": "table", "tableStyle": "Table Grid", "rows": [[spec["label"]], [text]]})
        roles.append({"role": "field"})
    return field_count, stats


def _flow_blocks(blocks, roles, fields, sections, header, title, subject):
    """整段版：字段名加粗成段，正文按行成段 —— 给习惯「读一篇」而不是「看一张表」的老师。"""
    blocks.append({
        "type": "paragraph",
        "align": "center",
        "runs": [{"text": _heading_text(title, subject), "bold": True, "sizePt": 16}],
    })
    blocks.append({
        "type": "table",
        "tableStyle": "Table Grid",
        "rows": [
            # 课题独占一行（课题最长，半宽格子怎么缩字号都放不下），值和框线在 _polish 里跨三列
            ["课题", header["title"], "", ""],
            ["课型", header["lessonType"], "课时", header["periods"]],
            ["授课教师", header["teacher"], "班级", header["className"]],
        ],
    })
    roles.append({"role": "header"})

    field_count = 0
    for spec in FIELDS:
        key = spec["key"]
        if key in HEADER_KEYS:
            continue
        if key == BOARD_KEY:
            blocks.append({"type": "table", "tableStyle": "Table Grid", "rows": [["板书设计"]],
                           "noRowSplit": True})
            roles.append({"role": "board", "lines": _lines(fields.get(BOARD_KEY))})
            continue
        text = str(fields.get(key) or "").strip()
        if not text:
            continue
        field_count += 1
        # keepWithNext：字段名不能被留在页底、内容翻到下一页（检查时最容易被挑「排版不规范」）
        blocks.append({"type": "paragraph", "keepWithNext": True,
                       "runs": [{"text": spec["label"], "bold": True}]})
        if key == PROCEDURE_KEY:
            for section in sections:
                if section["title"]:
                    caption = "（" + section["title"] + "）"
                    times = _clean_times(section["times"])
                    if times:
                        caption += "　" + "，".join(times)
                    blocks.append({"type": "paragraph", "keepWithNext": True,
                                   "runs": [{"text": caption, "bold": True}]})
                for line in section["lines"]:
                    blocks.append({"type": "paragraph", "runs": [{"text": line}]})
            continue
        for line in _lines(text):
            blocks.append({"type": "paragraph", "runs": [{"text": line}]})
    # 整段版没有 4 列表，也就没有「学生活动格」这一说：活动计数全 0，省得面板误以为有留白
    return field_count, {"rows": len(sections), "teacher": (0, 0),
                         "student": (0, 0), "time": (0, 0)}


def _set_widths(table, widths):
    """把列宽写死。

    build_docx 的表格块只能均分列宽，而教案表的列宽是内容决定的：
    「环节」窄、「教师/学生活动」宽、「时间」最窄。均分会把时间列撑得老宽、
    把活动内容挤成竖排 —— 一眼就能看出不是人排的。
    """
    total = int(sum(widths))
    tbl = table._tbl
    tbl_pr = tbl.tblPr
    tbl_w = tbl_pr.find(qn("w:tblW"))
    if tbl_w is None:
        tbl_w = tbl_pr.makeelement(qn("w:tblW"), {})
        tbl_pr.append(tbl_w)
    tbl_w.set(qn("w:type"), "dxa")
    tbl_w.set(qn("w:w"), str(total))
    grid = tbl.tblGrid
    if grid is not None:
        for column, width in zip(grid.findall(qn("w:gridCol")), widths):
            column.set(qn("w:w"), str(int(width)))
    for row in table.rows:
        for index, cell in enumerate(row.cells):
            if index >= len(widths):
                break
            try:
                cell.width = Twips(int(widths[index]))
            except Exception:
                break


def _text_ems(text):
    """粗估一段字的宽度，单位是「一个汉字的宽度」。

    ASCII 按 0.55 个字算：宋体/Times 里的英文都是半角，但 W、m、空格会把均值撑上去，
    按 0.5 估会在边界上判成「放得下」，实际却折行（12pt 的课题名就是这样折的）。
    """
    total = 0.0
    for char in str(text or ""):
        total += 1.0 if ord(char) > 0x2E80 else 0.55
    return total


def _fit_width_font(cell, width_twips, max_pt, min_pt):
    """按格子宽度定字号，让格子里的字尽量落在一行。

    Word 的表格没有「自动缩字」，只能自己算：可用宽度减掉左右内边距，
    除以估算字宽就得到能放下的字号；缩到 min_pt 还放不下的超长值（三四行的课题）
    只好让它折行 —— 再缩就打不出来了。
    """
    ems = _text_ems("".join(paragraph.text for paragraph in cell.paragraphs))
    if ems <= 0:
        return
    available = width_twips / 20.0 - 11.0  # 左右内边距各 108 twips ≈ 5.4pt
    size = max(min_pt, min(max_pt, available / ems))
    size = round(size * 2) / 2.0  # 半磅一档，字号跳跃不至于乱七八糟
    for paragraph in cell.paragraphs:
        for run in paragraph.runs:
            run.font.size = Pt(size)


def _bold_cell(cell, size_pt=None):
    for paragraph in cell.paragraphs:
        for run in paragraph.runs:
            run.bold = True
            if size_pt:
                run.font.size = Pt(size_pt)


def _fill_board(cell, lines):
    """板书格子：第一行「板书设计」加粗，下面留白给手写。

    格子里的文字要分段写 —— 一个大段落里混着加粗标签和正文，Word 只能整段加粗，
    结果整块板书都变粗，老师手写上去的笔迹和打印字分不开。
    """
    cell.text = ""
    label = cell.paragraphs[0]
    # cell.text 会留下一个空 run；不先清掉，加粗就落在这个空 run 上，标签仍然是细的
    for run in list(label.runs):
        run._r.getparent().remove(run._r)
    run = label.add_run("板书设计")
    run.bold = True
    run.font.size = Pt(11)
    for line in lines:
        cell.add_paragraph().add_run(line)


def _write_footer(document, text):
    """写页脚并用占位符换成真页码域（{{page}}/{{pages}} → w:fldSimple）。"""
    footer = document.sections[-1].footer
    footer.is_linked_to_previous = False  # 不先断开，footer 部件不会被创建
    paragraph = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
    # 必须连 fldSimple（页码域）一起清：只清 w:r 会把旧页码域留在段首，
    # 成品页脚就变成「12第 1 页 共 2 页」——两个数字是上一版页脚的残留
    for child in list(paragraph._p):
        if child.tag != qn("w:pPr"):
            paragraph._p.remove(child)
    run = paragraph.add_run(text)
    run.font.size = Pt(9)
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    fill_document(document, {}, include_body=False)


def _polish(path, roles, footer=None):
    """表格的格式只能在这一步补：build_docx 的表格块只收纯文字，不管加粗和列宽。"""
    document = docx.Document(path)
    tables = document.tables
    width = column_width_twips(document.sections[-1]) or 8300
    for index, role in enumerate(roles):
        if index >= len(tables):
            break
        table = tables[index]
        kind = role.get("role")
        if kind == "header":
            # 列宽写死成 标签|值|标签|值：均分会把「课时」这种短值的格子挤窄，值就折到第二行，
            # 看着像老师没填完 —— 抬头表最忌讳这个
            label = 1300
            value = (width - label * 2) // 2
            widths = [label, value, label, width - label * 2 - value]
            _set_widths(table, widths)
            for row_index in range(len(table.rows)):
                row = table.rows[row_index]
                row.height = Cm(0.95)
                row.height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST
                for column_index in range(len(widths)):
                    cell = table.cell(row_index, column_index)
                    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                    if column_index in (0, 2):
                        _bold_cell(cell)
                        _fit_width_font(cell, widths[column_index], 11.0, 9.0)
                    else:
                        # 值是 12pt 的正文尺寸，长一点就折行；这里按格宽缩到一行放得下
                        _fit_width_font(cell, widths[column_index], 10.5, 8.0)
            # 课题那一行的值横跨后三列：课题名一句英文加前缀就折两行，抬头表一折行就像没填完
            wide = width - label
            merged = table.cell(0, 1).merge(table.cell(0, len(table.columns) - 1))
            merged.width = Twips(wide)
            for paragraph in list(merged.paragraphs[1:]):
                if not paragraph.text.strip():
                    paragraph._p.getparent().remove(paragraph._p)
            _fit_width_font(merged, wide, 10.5, 8.0)
        elif kind == "field":
            _set_widths(table, [width])
            _bold_cell(table.cell(0, 0), size_pt=11)
            table.rows[0].height = Cm(0.9)
            table.rows[0].height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST
        elif kind == "procedure":
            step = 1400
            clock = 1100
            half = max(800, (width - step - clock) // 2)
            _set_widths(table, [step, half, width - step - clock - half, clock])
            for column_index in range(len(table.columns)):
                _bold_cell(table.cell(0, column_index), size_pt=10.5)
            for row_index in range(1, len(table.rows)):
                _bold_cell(table.cell(row_index, 0), size_pt=10.5)
            if role.get("merged") and len(table.rows) > 1:
                merged = table.cell(1, 0).merge(table.cell(1, len(table.columns) - 1))
                for paragraph in list(merged.paragraphs[1:]):
                    if not paragraph.text.strip():
                        paragraph._p.getparent().remove(paragraph._p)
        elif kind == "board":
            _set_widths(table, [width])
            lines = role.get("lines") or []
            _fill_board(table.cell(0, 0), lines)
            table.rows[0].height = Cm(4.0 if lines else 3.2)
            table.rows[0].height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST
    # 说明小字排成灰的：print 出来是「注」，不会和教案正文抢
    for paragraph in document.paragraphs:
        if paragraph.text.strip() == NOTE_BLANK:
            for run in paragraph.runs:
                run.font.color.rgb = RGBColor(0x80, 0x80, 0x80)
                run.font.size = Pt(9)
    if footer:
        _write_footer(document, footer)
    document.save(path)


def _require_template(template):
    """先自己把模板打开一次：坏文件给中文提示，别把 python-docx 的英文异常甩给老师。"""
    full = str(template)
    require(os.path.isfile(full), "file-not-found", "找不到这个模板文件：" + full,
            "确认模板没被移走，或重新导出一份 .docx 模板")
    try:
        docx.Document(full)
    except Exception:
        raise EngineError("template-broken", "这份模板打不开（不是 .docx，或者文件已损坏）：" + full,
                          "用 Word 打开它另存为 .docx，再试一次")
    return full


def build_plan(payload):
    """payload = {text, title?, layout?("form"|"flow"), teacher?, className?, subject?, periods?,
                  page?, baseTemplate?, outputPath?, overwrite?} → 教案 docx

    返回 {outputPath, title, layout, fieldCount, sections, warnings, bytes, table}：
    - fieldCount 是正文里排出来的字段块数量（不含顶部信息表；板书框固定出，也不算在内）
    - sections 是教学过程的环节标题，老师对着它核「导入/新授/操练」有没有漏
    - table 是教学过程表的统计：{rows, teacher:{filled,blank}, student:{filled,blank},
      time:{filled,blank}}。学生活动/时间留白是故意的（不瞎编学生的实际反应），
      面板拿这个计数就能解释「整列空白不是没生成」；整段版（flow）没有这张表，计数全 0
    - 顶部信息表是三行：课题|课型 / 授课教师|课时 / 班级（值跨三列）
    """
    payload = payload or {}
    parsed = _read_plan(payload.get("text"))
    fields = dict(parsed["fields"])
    sections = parsed["sections"]
    warnings = list(parsed["warnings"])

    layout = str(payload.get("layout") or "form").strip().lower()
    if layout not in ("form", "flow"):
        raise EngineError("unknown-layout", "不认识的版式：" + str(payload.get("layout")),
                          "可用：form（表格版，迎检用这个）、flow（整段版）")

    # 面板上手填的这几项覆盖解析结果：老师改了名字之后，成品必须按他改的印
    for key, value in (("teacher", payload.get("teacher")),
                       ("className", payload.get("className")),
                       ("periods", payload.get("periods"))):
        text = str(value or "").strip()
        if text:
            fields[key] = text

    title = _pick_title(payload, fields)
    subject = str(payload.get("subject") or "").strip()
    header = {
        "title": _one_line(title or fields.get("title")),
        "lessonType": _one_line(fields.get("lessonType")),
        "teacher": _one_line(fields.get("teacher")),
        "className": _one_line(fields.get("className")),
        "periods": _one_line(fields.get("periods")),
    }

    blocks = []
    roles = []
    if layout == "form":
        field_count, stats = _form_blocks(blocks, roles, fields, sections, header, title, subject)
    else:
        field_count, stats = _flow_blocks(blocks, roles, fields, sections, header, title, subject)

    if not sections or all(not section["title"] for section in sections):
        if fields.get(PROCEDURE_KEY):
            warnings.append({
                "code": "no-sections",
                "message": "教学过程没分出环节（导入/新授/操练…），已整段排成一行；"
                           "把每个环节另起一行写成「导入（5 分钟）」「（一）导入」"
                           "或「环节一：导入」，就能一格一个环节",
            })
    if not fields.get(PROCEDURE_KEY):
        warnings.append({
            "code": "no-procedure",
            "message": "没找到「教学过程」——这块是教案最要紧的一页，检查一下是不是漏粘了",
        })
    if stats["rows"] and not stats["student"][0] and stats["student"][1]:
        # 只在这种「整列都没写」的时候发：个别格子留白是正常的，条条报警反而让老师忽略真问题
        warnings.append({
            "code": "student-column-blank",
            "message": "「学生活动」整列都没写 —— 不是漏生成：这一列表头已标成「老师自填」，"
                       "格子下面还有一行说明，你按本班实际补几句或直接手写就行",
        })
    if field_count == 0:
        warnings.append({
            "code": "empty-plan",
            "message": "除了标题信息，一个正文字段都没认出来：每块另起一行写「教学目标：…」这种"
                       "「字段名：内容」，就能认出来",
        })
    for spec in FIELDS:
        key = spec["key"]
        if key in HEADER_KEYS or key == BOARD_KEY:
            continue
        if key in fields and not str(fields[key]).strip():
            warnings.append({
                "code": "empty-field",
                "message": "「" + spec["label"] + "」只有标题没有内容，成品里先不印这一块；"
                           "想留手写位置就随便写一句进去",
            })

    template = payload.get("baseTemplate")
    if template:
        template = _require_template(template)
    output = payload.get("outputPath") or default_output("教案", title, ".docx")

    spec = {
        "outputPath": output,
        "blocks": blocks,
        "values": {key: value for key, value in
                   (("title", title), ("teacher", header["teacher"]),
                    ("className", header["className"]), ("subject", subject)) if value},
        "properties": {"title": title, "author": header["teacher"] or "", "subject": "教案"},
        "overwrite": bool(payload.get("overwrite", False)),
    }
    if template:
        spec["baseTemplate"] = template
        # 模板自己的纸张/页边距/页眉页脚都是学校定的，不传 page 就不动它
        if payload.get("page"):
            spec["page"] = payload["page"]
    else:
        spec["page"] = payload.get("page") or {"size": "A4", "orientation": "portrait"}
        # 不要卷子版默认页脚（「出卷老师」），教案页脚在 _polish 里自己写
        spec["defaultFooter"] = False

    built = build_docx(spec)
    _polish(built["outputPath"], roles, footer=None if template else PLAN_FOOTER)
    warnings.extend(built.get("warnings") or [])

    return {
        "outputPath": built["outputPath"],
        "title": title,
        "layout": layout,
        "fieldCount": field_count,
        "sections": [section["title"] for section in sections if section["title"]],
        "warnings": warnings,
        "bytes": os.path.getsize(built["outputPath"]),
        # 教学过程表的三列各填了几格：面板靠它说清「留白是留给你填的，不是没生成」
        "table": {
            "rows": stats["rows"],
            "teacher": {"filled": stats["teacher"][0], "blank": stats["teacher"][1]},
            "student": {"filled": stats["student"][0], "blank": stats["student"][1]},
            "time": {"filled": stats["time"][0], "blank": stats["time"][1]},
        },
    }

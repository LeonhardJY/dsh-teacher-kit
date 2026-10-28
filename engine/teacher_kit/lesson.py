"""课件：把老师手里那段"上课要讲的东西"变成能直接投的 pptx。

管的不多：结构（封面 → 目标 → 内容 → 小结 → 作业）、不把要点挤成一坨、
版式用老师自己的学校模板。**不生成教学内容**——内容还是老师自己的，
这里只负责"排成一页一页、别丢要点、别一页塞 20 行"。

输入是老师真实会粘贴的样子：带【】的小标题、markdown #、一、二、三、、
光秃秃的一行「学习目标」/「New words」/「Practice」（最多人这么写），
或者干脆就是几段不带任何标记的话。
"""

import os
import re

from .errors import require
from .paths import default_output
from .pptx_build import build_pptx

RE_BRACKET_HEAD = re.compile(r"^\s*[【\[]\s*([^】\]]{1,20})\s*[】\]]\s*(.*)$")
RE_MD_HEAD = re.compile(r"^\s*(#{1,6})\s*(.+?)\s*#*\s*$")
RE_CN_HEAD = re.compile(r"^\s*([一二三四五六七八九十]{1,3})\s*[、.．]\s*(.{1,24})$")
# "1. xxx" / "1、xxx" / "1) xxx" / "1）xxx"：先一律当"编号行"，是不是小标题
# 由 _resolve_numeric_runs 按上下文判断（连续编号是有序列表，不是 N 个小节）。
RE_NUM_ITEM = re.compile(r"^\s*(\d{1,2})\s*([、.．)）])\s*(.*)$")
RE_ARABIC_LIST = re.compile(r"^\s*(\d{1,2})\s*[、.．)）]\s*(.+)$")
RE_BULLET = re.compile(r"^\s*[-–—·•*+]\s*(.+)$")
RE_SPEAK = re.compile(r"^\s*(?:讲|讲稿|讲解|备注)\s*[:：]\s*(.*)$")

# 这些标题本身值得单独一页，其余标题按普通内容页处理
KEY_SECTIONS = ("目标", "重点", "难点", "小结", "总结", "作业", "练习", "操练", "板书")
MAX_BULLETS = 6
MAX_SLIDE_CHARS = 300
# 一条要点在投影上的上限。中文 26pt 一行约 30 字，60 字 = 两行；
# 更长的整段（课文原文/英文段落）按句子、逗号、空白切开，绝不整段塞一页。
MAX_BULLET_CHARS = 60
MAX_NUM_HEAD_CHARS = 24
NUM_HEAD_STYLES = (".", "．", "、")
LINE_END_PUNCT = ("。", "！", "？", "?", "!", "；", ";", ".", "，", ",")

# ── 裸小标题（评审 P1-8 / P1-4）────────────────────────────────────────
# 老师写课件很少加【】或 #：一行「学习目标」、一行「New words」，下面才是要点。
# 以前这些行全被当成正文要点，6 页全叫「课题（续）」，投影上第一条就是「• 学习目标」。
MAX_BARE_HEAD_CHARS = 12
RE_HEAD_COLON = re.compile(r"^\s*([^:：]{1,24}?)\s*[:：]\s*(.*)$")
RE_LATIN_TOKEN = re.compile(r"[A-Za-z][A-Za-z'’\-]*")
# 一眼就是「这一页讲什么」的环节/教学词。整串相等才算 ——「复习单词」是正文，不是小节名。
# Practice / Production / Presentation 留在这里：英语课的 P-P-P 环节名，不是一句活动指令。
SECTION_HEAD_WORDS = (
    "学习目标", "教学目标", "知识目标", "能力目标", "情感目标", "重点", "难点", "重点难点",
    "教学重点", "教学难点", "重点词汇", "重点句型", "重点短语", "重点句子",
    "生词", "新词", "单词", "词汇", "句型", "语法", "知识点", "语言点", "课文理解",
    "导入", "新课导入", "导入新课", "热身", "复习", "复习导入", "预习",
    "新授", "讲授", "精讲", "操练", "练习", "巩固", "巩固练习", "当堂检测", "检测", "反馈",
    "小结", "总结", "课堂小结", "归纳", "作业", "作业布置", "布置作业", "板书", "板书设计",
    "拓展", "延伸", "拓展延伸", "活动", "任务", "讨论", "探究", "展示", "游戏",
    "听说", "读写", "阅读", "写作", "听力", "口语", "课堂练习", "课后作业",
    "warm-up", "warm up", "lead-in", "lead in", "presentation", "practice",
    "production", "summary", "homework", "revision", "review", "greeting",
    "reading", "writing", "listening", "speaking", "new words", "key words",
    "key sentences", "key points", "phrases", "vocabulary", "grammar",
    "exercise", "exercises", "consolidation", "conclusion", "objectives",
    "aims", "goals", "self-check", "self check", "pre-reading", "while-reading",
    "post-reading", "section a", "section b",
)
# 课堂活动名（评审 P1-4）：写在「小标题」的位置上，也只是上一节的活动要点。
# 把它们升成页，就会多出一页只有「操练」两个字、正文全在下一页的空页
# （老师原文「- Pair work：描述你的同桌」，Pair work 曾被当成一节标题）。
# 中文的「游戏/活动/任务」留在上面 —— 教案里「二、游戏」确实当小节用；
# 升页的词只认「这一页讲什么」，不认「这一节做什么」。
ACTIVITY_HEAD_WORDS = (
    "pair work", "pairwork", "paired work", "pair activity", "pair works",
    "group work", "groupwork", "group activity", "group discussion",
    "guessing game", "guessing games", "word game", "role play", "role-play", "roleplay",
    "game", "games", "activity", "activities", "task", "tasks",
    "discussion", "free talk",
)
BARE_HEAD_WORDS = SECTION_HEAD_WORDS + ACTIVITY_HEAD_WORDS   # 整张词表（旧名保留）
SECTION_HEAD_SET = {re.sub(r"[\s\u3000]+", " ", word).strip().casefold() for word in SECTION_HEAD_WORDS}
ACTIVITY_HEAD_SET = {re.sub(r"[\s\u3000]+", " ", word).strip().casefold() for word in ACTIVITY_HEAD_WORDS}
BARE_HEAD_SET = SECTION_HEAD_SET | ACTIVITY_HEAD_SET


def _looks_like_heading(line):
    """短、不带句号、不是一串数字的量词——才当标题。老师正文里"一、"开头的句子不能误伤。

    数字编号（"1. xxx"）不在这里判：它可能是有序列表的第 1 条，
    交给 _resolve_numeric_runs 看上下文。
    """
    stripped = line.strip()
    if not stripped or stripped.endswith("。") or stripped.endswith("？") or stripped.endswith("?"):
        return None
    bracket = RE_BRACKET_HEAD.match(stripped)
    if bracket:
        return (bracket.group(1).strip(), bracket.group(2).strip())
    md = RE_MD_HEAD.match(stripped)
    if md:
        return (md.group(2).strip(), "")
    cn = RE_CN_HEAD.match(stripped)
    if cn:
        return (cn.group(2).strip(), "")
    return None


def _norm_head(text):
    return re.sub(r"[\s\u3000]+", " ", str(text or "")).strip().casefold()


def _is_capitalized_phrase(text):
    """「New Words」「Key Sentences」这种首字母全大写的英文短语。

    单个英文词不在这里认（Practice 之类走教学词表）——
    否则课文里的人名（Tom）、地名单独一行也会变成一页标题。
    """
    if any(ord(char) > 0x2E80 for char in text):
        return False
    tokens = RE_LATIN_TOKEN.findall(text)
    if not 2 <= len(tokens) <= 4:
        return False
    return all(token[0].isupper() for token in tokens)


def _bare_heading(line):
    """一行裸小标题 → (标题, 行内正文)；不是小标题返回 None。

    判据（评审 P1-8）：≤12 字、不带句末标点、不带数字、不带编号，
    整串就在「环节/教学词表」（SECTION_HEAD_SET）里，或者是首字母全大写的英文短语。
    带数字的一律排除：「认识 6 个新单词」是正文 —— 认错一行，投影上就多一页空标题。
    「• 学习目标」这种带项目符号的也认：它本身就是小标题，不能留在正文里当第一条要点。

    课堂活动名（评审 P1-4）一律不升页：「pair work / group work / guessing game」
    这类只是「这一节做什么」，写成「- Pair work：描述你的同桌」时更是一条要点；
    认成标题就会把上一节（操练）的正文全搬到新页，留下光秃秃一页标题。
    这条判断放在大写字短语之前 ——「Pair Work」「Group Work」同样不算标题。
    """
    text = str(line or "").strip()
    if not text:
        return None
    bullet = RE_BULLET.match(text)
    if bullet:
        text = bullet.group(1).strip()
    match = RE_HEAD_COLON.match(text)
    head, inline = (match.group(1).strip(), match.group(2).strip()) if match else (text, "")
    if not head or len(head) > MAX_BARE_HEAD_CHARS:
        return None
    if any(char in head for char in "。！？!?；;，,、："):
        return None
    if any(char.isdigit() for char in head):
        return None
    norm = _norm_head(head)
    if norm in ACTIVITY_HEAD_SET:
        return None
    if norm in SECTION_HEAD_SET:
        return head, inline
    if _is_capitalized_phrase(head):
        return head, inline
    return None


def _body_follows(lines, index):
    """这一行下面真有正文吗？只写了个小标题、后面空到结尾的话，认出来就是一张空页。"""
    for raw in lines[index + 1:]:
        text = str(raw or "").strip()
        if not text or RE_SPEAK.match(text):
            continue
        if _looks_like_heading(text) or RE_NUM_ITEM.match(text):
            return True
        if _bare_heading(text):
            continue  # 紧跟的也是个小标题：接着往后找正文
        return True
    return False


def _wrap(text, limit=MAX_BULLET_CHARS):
    """超长的一段没有句号可切时，按空白/定长硬切——宁可两条短句，也别一条 500 字。"""
    text = (text or "").strip()
    if not text:
        return []
    if len(text) <= limit:
        return [text]
    pieces = []
    buffer = ""
    for token in re.split(r"(\s+)", text):
        if not token:
            continue
        if len(buffer) + len(token) <= limit:
            buffer += token
            continue
        if buffer.strip():
            pieces.append(buffer.strip())
        buffer = token.strip()
        while len(buffer) > limit:
            pieces.append(buffer[:limit])
            buffer = buffer[limit:]
    if buffer.strip():
        pieces.append(buffer.strip())
    return pieces or [text]


def _split_at_pauses(text, limit=MAX_BULLET_CHARS):
    """按逗号/顿号/冒号这类句内停顿先切，再按长度兜底。"""
    text = (text or "").strip()
    if len(text) <= limit:
        return [text] if text else []
    clauses = [clause.strip() for clause in re.split(r"(?<=[，,、：:])", text) if clause.strip()]
    pieces = []
    buffer = ""
    for clause in clauses:
        if len(clause) > limit:
            if buffer:
                pieces.append(buffer)
                buffer = ""
            pieces.extend(_wrap(clause, limit))
            continue
        if buffer and len(buffer) + len(clause) > limit:
            pieces.append(buffer)
            buffer = clause
        else:
            buffer += clause
    if buffer:
        pieces.append(buffer)
    return pieces or _wrap(text, limit)


def _sentences(text):
    """把一段话切成能上投影的短句（一句话一个要点）。

    短行不切——"能用 What does he look like? 询问外貌。"这类中英混排的一句，
    在 ? 后切开只会得到两个半截。整行超过 30 字才拆，拆也只认中文句末标点；
    单句仍然超过一条要点的上限时，再按句内停顿/长度切成多条。
    """
    text = re.sub(r"[ \t]+", " ", (text or "").strip())
    if not text:
        return []
    if len(text) <= 30:
        return [text]
    parts = re.split(r"(?<=[。！？；])", text)
    pieces = []
    for part in parts:
        part = part.strip()
        if not part:
            continue
        pieces.extend(_split_at_pauses(part))
    return pieces or [text]


def _bullets_from_line(line):
    bullet = RE_BULLET.match(line)
    if bullet:
        return [part for part in _sentences(bullet.group(1)) if part]
    arabic = RE_ARABIC_LIST.match(line)
    if arabic and arabic.group(2).strip():
        # "1. 掌握 5 个单词。" 这种编号句：编号由分页/排序负责，投影上不要顶着 1.
        return [part for part in _sentences(arabic.group(2)) if part] or [arabic.group(2).strip()]
    return _sentences(line)


def _numeric_item(line):
    match = RE_NUM_ITEM.match(line)
    if not match:
        return None
    text = match.group(3).strip()
    if not text:
        return None
    return {"kind": "num", "num": int(match.group(1)), "style": match.group(2), "text": text}


def _num_is_heading(records, index):
    """单独一条编号行只有"像标题"且后面确实有正文时才算小节。

    连续/密集的编号（S1-2：1..7 全是要点）会被 _resolve_numeric_runs 整段降级；
    这里处理孤立的一条：短、不带句末标点、下面真有正文。否则当正文要点，
    免得产出一页只有标题、正文空白。
    """
    record = records[index]
    if record.get("style") not in NUM_HEAD_STYLES:
        return False
    text = record["text"]
    if len(text) > MAX_NUM_HEAD_CHARS or text.endswith(LINE_END_PUNCT):
        return False
    for follow in records[index + 1:]:
        if follow["kind"] == "speak":
            continue
        return follow["kind"] == "body"
    return False


def _resolve_numeric_runs(records):
    """把数字编号行分成两类：小节标题 / 有序列表要点。

    判据是"上下文"而不是"这一行短不短"：
    - 连续两条以上的编号行（中间只夹讲稿）→ 一整个有序列表，全部当正文要点；
    - 孤立的一条编号行 → 只有短、无句末标点、后面紧跟正文时才算小节标题。
    这样 "1. 学习目标 / 2. 重点词汇 / 3. 句型操练"（每节下面有要点）仍是三页，
    而 "1. ... 2. ... 3. ..." 的备课清单不会变成 N 个空页。
    """
    resolved = []
    total = len(records)
    index = 0
    while index < total:
        record = records[index]
        if record["kind"] != "num":
            resolved.append(record)
            index += 1
            continue
        run = []
        cursor = index
        while cursor < total:
            kind = records[cursor]["kind"]
            if kind == "num":
                run.append(cursor)
                cursor += 1
            elif kind == "speak":
                cursor += 1
            else:
                break
        heading = len(run) == 1 and _num_is_heading(records, run[0])
        for position in range(index, cursor):
            item = records[position]
            if item["kind"] != "num":
                resolved.append(item)
            elif heading:
                resolved.append({"kind": "heading", "title": item["text"], "residual": ""})
            else:
                resolved.append({"kind": "body", "text": item["text"]})
        index = cursor
    return resolved


def _classify(lines):
    """先把每一行分类（讲稿/编号/标题/正文），再做编号消歧。

    小标题四种写法都认：【小标题】/ # 小标题 / 一、小标题，以及老师最常写的裸小标题
    （「学习目标」「New words」「Practice」）—— 裸小标题要往后看一眼，确认下面真有正文。
    """
    text_lines = [str(raw or "").strip() for raw in lines]
    records = []
    for index, line in enumerate(text_lines):
        if not line:
            continue
        speak = RE_SPEAK.match(line)
        if speak:
            records.append({"kind": "speak", "text": speak.group(1).strip()})
            continue
        numeric = _numeric_item(line)
        if numeric:
            records.append(numeric)
            continue
        heading = _looks_like_heading(line)
        if heading:
            records.append({"kind": "heading", "title": heading[0], "residual": heading[1]})
            continue
        bare = _bare_heading(line)
        if bare and _body_follows(text_lines, index):
            records.append({"kind": "heading", "title": bare[0], "residual": bare[1]})
            continue
        records.append({"kind": "body", "text": line})
    return _resolve_numeric_runs(records)


def _bullet_text(item):
    return (item if isinstance(item, str) else (item.get("text") or "")).strip()


def _dedupe_bullets(title, cover_title, bullets):
    """页标题/封面标题又在正文第一行出现时去掉，别让投影上第一行就是标题。"""
    headings = {str(title or "").strip(), str(cover_title or "").strip()}
    seen = set()
    kept = []
    for item in bullets:
        text = _bullet_text(item)
        if not text or text in headings or text in seen:
            continue
        seen.add(text)
        kept.append(item)
    return kept or list(bullets)


def _compact(text):
    return re.sub(r"[\s，。、·:：\-—（）()《》\"']+", "", str(text or "")).lower()


def _similar_text(left, right):
    """两个短标题是不是同一件事（「Unit 6 食物」对「Unit 6 Food 食物」）。

    字面相等只是最省事的一种情况：老师的封面写在 title 里、大纲首行又用另一种写法
    写了一遍课题，以前只认字面相等，于是第 2 页又出一张同标题封面。
    """
    left, right = _compact(left), _compact(right)
    if not left or not right:
        return False
    if left == right:
        return True
    if left in right or right in left:
        shorter, longer = min(len(left), len(right)), max(len(left), len(right))
        return shorter / float(longer) >= 0.6
    left_chars, right_chars = set(left), set(right)
    return len(left_chars & right_chars) / float(len(left_chars | right_chars)) >= 0.75


def _drop_duplicate_cover(blocks, cover_title, subtitle):
    """「首行就是课题」的大纲会多出一页：标题与封面相同、正文只有课题原文。

    判据：标题等于（或近似）封面标题、要点 ≤1 条，且这条要点去掉标点后包含于
    「封面标题 + 副标题」之中（或反过来）。命中就把整块并进封面，不单独成页。
    近似只对**第一块**放宽：中间某一页标题碰巧像封面，那是真内容，不能丢。
    """
    cover = _compact(cover_title) + _compact(subtitle)
    kept = []
    dropped = 0
    for index, (title, bullets, notes) in enumerate(blocks):
        first = index == 0
        same_title = title == cover_title or (first and _similar_text(title, cover_title))
        if same_title and len(bullets) <= 1 and cover:
            same = True
            for bullet in bullets:
                compact = _compact(bullet)
                near = first and (_similar_text(bullet, cover_title) or _similar_text(bullet, cover))
                if not compact or not (compact in cover or cover in compact or near):
                    same = False
                    break
            if same:
                dropped += 1
                continue
        kept.append((title, bullets, notes))
    return kept, dropped


def _merge_blocks(blocks):
    """同一标题连着出现两次（老师把小标题写重了）就并成一块。

    并块而不是各出一页：投影上两页同名更难认，也会逼出「标题（续）」；
    「（续）」只留给真的要点超量、被 _chunk 拆页的情况。
    """
    merged = []
    for title, bullets, notes in blocks:
        if merged and merged[-1][0] == title:
            previous = merged[-1]
            merged[-1] = (previous[0], previous[1] + list(bullets), previous[2] or notes)
        else:
            merged.append((title, list(bullets), notes))
    return merged


def _absorb_title_only(blocks):
    """只剩标题、一条正文都没有的块不许单独成页：把标题并回上一页（评审 P1-4）。

    投影上留一页光秃秃的小标题，老师只能站在原地翻页；并回上一页后，
    这个标题至少还以要点的形式留在页面上——页数只会变少，一条内容都不丢。
    前面只有封面（没有可并的正文页）时保持原样，交给 _chunk 报 empty-slide：
    「真的只写了小标题」这种情况仍然要吵出来，不能静默。
    """
    kept = []
    for title, bullets, notes in blocks:
        if bullets:
            kept.append((title, list(bullets), notes))
            continue
        label = str(title or "").strip()
        target = None
        for position in range(len(kept) - 1, -1, -1):
            if kept[position][1]:
                target = position
                break
        if target is None:
            kept.append((title, [], notes))
            continue
        previous_title, previous_bullets, previous_notes = kept[target]
        merged = list(previous_bullets)
        if label and label != str(previous_title or "").strip() and label not in merged:
            merged.append(label)
        kept[target] = (previous_title, merged, previous_notes or notes)
    return kept


def _chunk(slides, title, bullets, notes, warnings):
    """一页放 6 条 / 300 字以内，超了就开续页——宁可多一页，也别让投影挤成一坨。"""
    if not bullets:
        warnings.append({
            "slide": len(slides) + 1,
            "code": "empty-slide",
            "message": "第 " + str(len(slides) + 1) + " 页「" + str(title) + "」只有标题没有正文——"
                       "检查一下这段是不是只写了小标题",
        })
        slides.append({"title": title, "bullets": [], "notes": notes})
        return
    part = 0
    index = 0
    while index < len(bullets):
        chunk = []
        chars = 0
        while index < len(bullets) and len(chunk) < MAX_BULLETS:
            item = bullets[index]
            text = _bullet_text(item)
            if chunk and chars + len(text) > MAX_SLIDE_CHARS:
                break
            chunk.append(item if isinstance(item, dict) else {"text": item, "level": 0})
            chars += len(text)
            index += 1
        page_title = title if part == 0 else title + "（续）"
        if part > 0:
            warnings.append({"slide": len(slides) + 1, "code": "split-slide",
                             "message": "「" + title + "」要点太多，已拆成第 " + str(part + 1) + " 页"})
        slides.append({"title": page_title, "bullets": chunk,
                       "notes": notes if part == 0 else None})
        part += 1
    _balance_last(slides, title)


def _balance_last(slides, title):
    """最后一页只剩 1 条、前一页却挤了 3 条以上时，挪过来 1-2 条。

    投影上「6 条 + 1 条孤零零一行」比「4 条 + 3 条」难看，老师也容易以为是漏了。
    """
    if len(slides) < 2:
        return
    last, previous = slides[-1], slides[-2]
    if not isinstance(last.get("bullets"), list) or not isinstance(previous.get("bullets"), list):
        return
    if previous.get("title") not in (title, title + "（续）"):
        return
    while len(last["bullets"]) < 2 and len(previous["bullets"]) >= 3:
        last["bullets"].insert(0, previous["bullets"].pop())


def plan_deck(payload):
    """payload = {text, title?, subtitle?, options?} → {slides, warnings, headings}"""
    text = payload.get("text")
    require(bool((text or "").strip()), "text-required", "没有收到课件内容（把这段课要讲的东西粘贴进来）", "按「标题 + 要点」分段粘：一行标题，下面几行是这一页要讲的点")
    options = payload.get("options") or {}
    lines = [line.rstrip() for line in text.replace("\r\n", "\n").replace("\r", "\n").split("\n")]

    cover_title = payload.get("title") or options.get("title") or "今天这节课"
    subtitle = payload.get("subtitle") or options.get("subtitle")
    if not subtitle:
        pieces = [options.get(key) for key in ("grade", "subject", "teacher")]
        subtitle = " · ".join(str(piece) for piece in pieces if piece) or None

    slides = [{"title": cover_title, "subtitle": subtitle, "layout": "title"}]
    warnings = []

    blocks = []          # [(标题, [要点])]
    current_title = None
    current_bullets = []
    current_notes = []

    def flush():
        nonlocal current_title, current_bullets, current_notes
        if current_title is None and not current_bullets:
            current_notes = []
            return
        blocks.append((current_title if current_title is not None else cover_title,
                       list(current_bullets), " ".join(current_notes).strip() or None))
        current_title = None
        current_bullets = []
        current_notes = []

    for record in _classify(lines):
        kind = record["kind"]
        if kind == "speak":
            current_notes.append(record["text"])
        elif kind == "heading":
            flush()
            current_title = record["title"] or cover_title
            residual = (record.get("residual") or "").strip()
            if residual:
                current_bullets.extend(_bullets_from_line(residual))
        else:
            current_bullets.extend(_bullets_from_line(record["text"]))

    flush()

    if not blocks:
        blocks.append((cover_title, [], None))

    blocks = _merge_blocks(blocks)
    blocks = _absorb_title_only(blocks)
    no_heading = all(_compact(title) == _compact(cover_title) for title, _bullets, _notes in blocks)
    blocks, cover_dropped = _drop_duplicate_cover(blocks, cover_title, subtitle)
    if cover_dropped:
        warnings.append({
            "code": "duplicate-cover",
            "message": "正文第一行就是课题（和封面标题重复），已并进封面，不再单独占一页空白",
        })
    blocks = [(title, _dedupe_bullets(title, cover_title, bullets), notes)
              for title, bullets, notes in blocks]
    # 没有小标题时，每一块标题都是自动填的封面标题 → 出稿后第 2 页顶着一行跟封面
    # 一模一样的标题（「Unit 6 食物」）。这种页改叫「课堂内容」，老师一眼知道是正文页。
    blocks = [(("课堂内容" if _compact(title) == _compact(cover_title) else title), bullets, notes)
              for title, bullets, notes in blocks]

    # 没认出任何标题时提示一句（分页由 _chunk 按 6 条 / 300 字负责，别把 20 条塞一页）
    if no_heading and len(blocks[0][1]) > MAX_BULLETS:
        warnings.append({"code": "no-heading",
                         "message": "没找到小标题，已按每页 " + str(MAX_BULLETS) + " 条自动分页——"
                                    "把「学习目标」「New words」这种小标题单独占一行，会排得更像课件"})

    # 正文里已经有「小结」「作业」小节时，options 里再给一份不该在末尾再排一页：
    # 以前是「…小结、作业、作业、小结」，最后两页还跟正文顺序相反。
    def _merge_option_into_blocks(key, word):
        value = options.get(key)
        if not value:
            return
        sentences = _sentences(str(value))
        if not sentences:
            return
        for index, (title, bullets, notes) in enumerate(blocks):
            if word in _compact(title):
                blocks[index] = (title, list(bullets) + sentences, notes)
                return
        _chunk(slides, word, sentences, None, warnings)

    _merge_option_into_blocks("summary", "小结")
    _merge_option_into_blocks("homework", "作业")

    for title, bullets, notes in blocks:
        _chunk(slides, title, bullets, notes, warnings)

    return {
        "slides": slides,
        "warnings": warnings,
        "outline": [{"title": slide.get("title"), "bullets": len(slide.get("bullets") or [])} for slide in slides],
    }


def build_deck(payload):
    """payload = {text, outputPath, title?, subtitle?, template?, options?{grade,subject,teacher,homework,summary,slideSize,notes}}"""
    options = payload.get("options") or {}
    planned = plan_deck(payload)
    output = payload.get("outputPath") or default_output(
        "课件", payload.get("title") or options.get("title") or "课件", ".pptx")

    # notes 有两个意思：true 只是「把讲稿写进备注」的开关，文本才是封面备注的内容。
    # 以前不区分，开关为真时封面备注第一行印成「True」。
    notes = options.get("notes")
    if isinstance(notes, (list, tuple)):
        notes = " ".join(str(item) for item in notes if item)
    if isinstance(notes, str) and notes.strip():
        planned["slides"][0]["notes"] = notes.strip()
    elif notes not in (None, False, True, "", 0):
        planned["warnings"].append({
            "code": "notes-ignored",
            "message": "封面备注只认文字（notes 传的是 " + type(notes).__name__
                       + "），这次没写上去；要写就传一段字符串",
        })

    spec = {
        "outputPath": output,
        "slides": planned["slides"],
        "titleSizePt": options.get("titleSizePt") or 36,
        "bodySizePt": options.get("bodySizePt") or 26,
        "properties": {
            "title": payload.get("title") or options.get("title") or "课件",
            "author": str(options.get("teacher") or ""),
        },
        "overwrite": bool(payload.get("overwrite", False)),
    }
    if payload.get("template"):
        spec["template"] = payload["template"]
        spec["clearTemplateSlides"] = options.get("clearTemplateSlides", True)
    else:
        spec["slideSize"] = options.get("slideSize") or "16:9"
    built = build_pptx(spec)

    return {
        "outputPath": built["outputPath"],
        "slides": built["slides"],
        "outline": planned["outline"],
        "slideSize": built["slideSize"],
        "template": os.path.abspath(payload["template"]) if payload.get("template") else None,
        "removedTemplateSlides": built.get("removedTemplateSlides", 0),
        "warnings": list(planned["warnings"]) + list(built.get("warnings") or []),
    }

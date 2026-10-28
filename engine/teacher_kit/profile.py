"""老师与学校的画像：这个助手「越用越像你们学校」的那份记忆。

为什么需要它
------------
同一个「出试卷」，王老师要 8 开两栏、页脚印「第 X 页 共 Y 页」、出卷人写自己名字；
李老师要 A3 一栏、不印班级。以前的答案是固定表单：填完就生成一套标准化东西，换谁都一样，
第二次用还得重填。画像把这些「说过一次的话」留下来，下次在对话里直接当默认值用，
生成完再问一句「记成你们学校默认吗」——由老师决定记不记，工具不偷偷改。

存哪儿
------
默认 ~/.dsh/teacher-kit/profile.json（TEACHER_KIT_PROFILE 可覆盖），放在插件自己的状态目录，
不混进「文档/教师助手」那堆成品里。写不进去也不影响任何生成能力：读不到就当没有画像。

结构
----
分节存，节名固定（school/paper/plan/deck/wechat/docflow/tasks），节内键名不预设
——老师与学校的样子太杂，硬编码字段只会漏。值只收字符串/数字/布尔，长度有上限。
"""

import copy
import json
import os
import re
import time

SECTIONS = ("school", "paper", "plan", "deck", "wechat", "docflow", "tasks")
MAX_VALUE_LEN = 200
MAX_HISTORY = 50
MAX_BYTES = 64 * 1024
# 键名允许中文：「设备编号」「校歌」这种老师顺手写出来的键不该被静默丢掉（以前回执 ok:true 却 saved 全空）
RE_KEY = re.compile(r"^[A-Za-z\u4e00-\u9fff][A-Za-z0-9_\u4e00-\u9fff]{0,31}$")

DEFAULTS = {
    "version": 1,
    "school": {},
    "paper": {},
    "plan": {},
    "deck": {},
    "wechat": {},
    "docflow": {},
    "tasks": {},
    "learned": [],
    "history": [],
}


def profile_path():
    """画像文件的位置：环境变量优先，其次插件状态目录，最后退到教师助手目录。"""
    override = os.environ.get("TEACHER_KIT_PROFILE")
    if override:
        return os.path.abspath(override)
    home = os.path.expanduser("~")
    primary = os.path.join(home, ".dsh", "teacher-kit", "profile.json")
    if os.path.isdir(os.path.dirname(primary)):
        return primary
    # 状态目录还没建（或这台机器没有 ~/.dsh）：落在老师看得见的目录里，至少别丢
    from .paths import teacher_home
    return os.path.join(teacher_home(), "profile.json")


def empty():
    return copy.deepcopy(DEFAULTS)


def _merge(base, patch):
    """深合并非空值：字典递归，列表/标量整体替换。None 表示"这一项我不知道"，跳过。"""
    if not isinstance(patch, dict):
        return base
    for key, value in patch.items():
        if value is None:
            continue
        if isinstance(value, dict):
            current = base.get(key)
            if not isinstance(current, dict):
                current = {}
            base[key] = _merge(current, value)
        else:
            base[key] = value
    return base


def _clean_section(name, patch):
    """只留下能存的东西：键名合法、值类型收敛、字符串去空白并截断。"""
    out = {}
    if not isinstance(patch, dict):
        return out
    for key, value in patch.items():
        if not isinstance(key, str) or not RE_KEY.match(key):
            continue
        if isinstance(value, bool):
            out[key] = value
        elif isinstance(value, (int, float)):
            out[key] = value
        elif isinstance(value, str):
            text = value.strip()
            if text:
                out[key] = text[:MAX_VALUE_LEN]
        elif isinstance(value, (list, tuple)):
            # 只收字符串列表（比如"这套卷子固定有哪几个大题"），别的类型一律丢掉
            items = [str(item).strip()[:MAX_VALUE_LEN] for item in value if str(item).strip()]
            if items:
                out[key] = items[:20]
    return out


def sanitize(patch):
    """把外部传进来的 patch 洗成可存的形状；不认识的节直接忽略（不报错）。"""
    clean = {}
    ignored = []
    if not isinstance(patch, dict):
        return clean, ignored
    for name, value in patch.items():
        if name not in SECTIONS:
            ignored.append(name)
            continue
        section = _clean_section(name, value)
        if section:
            clean[name] = section
    return clean, ignored


def load():
    """读画像：文件不存在 / 坏了都返回空画像，绝不抛错（生成不能被记忆拖垮）。"""
    path = profile_path()
    profile = empty()
    if not os.path.isfile(path):
        return profile, path, False
    try:
        with open(path, "r", encoding="utf-8") as handle:
            raw = json.load(handle)
    except Exception:
        # 画像文件坏了：留一份 .broken 供回看/恢复，别静默当"从没记过"把老师习惯丢掉
        try:
            with open(path, "r", encoding="utf-8", errors="replace") as handle:
                broken = handle.read()
            if broken.strip():
                with open(path + ".broken", "w", encoding="utf-8") as handle:
                    handle.write(broken)
        except Exception:
            pass
        return profile, path, False
    if not isinstance(raw, dict):
        return profile, path, False
    for name in SECTIONS:
        profile[name] = _clean_section(name, raw.get(name)) if isinstance(raw.get(name), dict) else {}
    for name in ("learned", "history"):
        items = raw.get(name)
        if isinstance(items, list):
            profile[name] = [item for item in items if isinstance(item, dict)][-MAX_HISTORY:]
    return profile, path, True


def save(profile):
    """原子写：先写临时文件再替换，中途断电不会留下半个 JSON。"""
    path = profile_path()
    parent = os.path.dirname(path)
    if parent and not os.path.isdir(parent):
        os.makedirs(parent, exist_ok=True)
    text = json.dumps(profile, ensure_ascii=False, indent=2)
    if len(text.encode("utf-8")) > MAX_BYTES:
        # 超了只可能是 history 堆太多：砍掉旧的再存
        profile["history"] = profile.get("history", [])[-10:]
        text = json.dumps(profile, ensure_ascii=False, indent=2)
    # 临时名必须唯一：两个并发 profile:set 共用同一 .tmp 时，
    # 实测会把 profile.json 写成非法 JSON（Extra data），下次 load 静默丢整份画像。
    temp = "%s.%s.%s.tmp" % (path, os.getpid(), time.time_ns())
    try:
        with open(temp, "w", encoding="utf-8") as handle:
            handle.write(text)
        os.replace(temp, path)
    finally:
        if os.path.exists(temp):
            try:
                os.remove(temp)
            except OSError:
                pass
    return path


def cn_ordinal(number, fallback=None):
    """1 → 一、10 → 十、12 → 十二、21 → 二十一。超过 99 用阿拉伯数字，绝不 IndexError。"""
    digits = "一二三四五六七八九"
    try:
        number = int(number)
    except (TypeError, ValueError):
        return str(fallback if fallback is not None else number)
    if number < 1 or number > 99:
        return str(number)
    if number <= 9:
        return digits[number - 1]
    if number == 10:
        return "十"
    tens, ones = divmod(number, 10)
    head = "十" if tens == 1 else digits[tens - 1] + "十"
    return head + (digits[ones - 1] if ones else "")


def _stamp():
    return time.strftime("%Y-%m-%d %H:%M")


def update(patch, note=None, task=None, learn=False):
    """把老师这次说的话并进画像；learn=True 时顺手记一条"学到了什么"。

    note 是给老师看的一句话（"这套卷子 8 开两栏，已经记住了"），
    写进 history 让人以后能回看"这工具是怎么变成现在这样的"。
    """
    profile, path, _exists = load()
    clean, ignored = sanitize(patch)
    _merge(profile, clean)
    entry = {"at": _stamp(), "task": str(task or ""), "note": str(note or "").strip()[:MAX_VALUE_LEN * 2]}
    if clean:
        entry["keys"] = sorted(name + "." + key for name, value in clean.items() for key in value)
    if learn and (clean or note):
        profile["learned"] = (profile.get("learned") or [])[-MAX_HISTORY + 1:] + [entry]
    profile["history"] = (profile.get("history") or [])[-(MAX_HISTORY - 1):] + [entry]
    profile["version"] = 1
    save(profile)
    return {"profile": profile, "path": path, "ignored": ignored, "saved": sorted(name for name in clean)}


def reset():
    """清空画像（只清记忆，不动任何已经生成的文件）。"""
    path = profile_path()
    profile = empty()
    save(profile)
    return {"profile": profile, "path": path}


PAPER_LABELS = {
    "8K-2": "8 开横向两栏",
    "8K-1": "8 开横向一栏",
    "A3-2": "A3 横向两栏",
    "A3-1": "A3 横向一栏",
    "16K-1": "16 开一栏",
    "A4-1": "A4 一栏",
}

PLAN_LAYOUTS = {"form": "表格式", "flow": "流程版"}

# 老师/模型嘴里的纸型大小 → 内部名（'8开' 和 '8K' 是一回事）
PAPER_SIZES = {"8开": "8K", "8K": "8K", "A3": "A3", "16开": "16K", "16K": "16K", "A4": "A4"}


def paper_label(key):
    if not key:
        return ""
    return PAPER_LABELS.get(str(key), str(key))


def paper_key_of(paper):
    """画像里记的纸型 → '8K-2' 这种内部名。

    画像允许两种记法：直接记 `key`（8K-2），或者记老师嘴里的
    `{size: '8K', orientation: 'landscape', columns: 2}`。后一种是老师自己报的，
    更自然，所以这里要能把它折算回内部名；认不出来就原样返回（别编一个）。
    """
    if isinstance(paper, str):
        paper = {"size": paper}
    if not isinstance(paper, dict):
        return ""
    key = str(paper.get("key") or "").strip()
    for candidate in (key, str(paper.get("size") or "").strip()):  # 两种记法都可能塞了内部名
        if candidate in PAPER_LABELS:
            return candidate
    raw = str(key or paper.get("size") or "").strip().upper().replace(" ", "")
    size = PAPER_SIZES.get(raw) or PAPER_SIZES.get(raw.replace("K", "开"))
    if not size:
        return key
    if size in ("8K", "A3"):
        orientation = str(paper.get("orientation") or "").lower()
        try:
            columns = int(paper.get("columns")) if paper.get("columns") not in (None, "") else None
        except (TypeError, ValueError):
            columns = None
        if orientation in ("portrait", "纵向") or columns == 1:
            return size + "-1"
        return size + "-2"
    return size + "-1"


def paper_page(paper):
    """画像里的纸型 → paper.resolve 认的 {size, orientation, columns}；认不出给 None。"""
    key = paper_key_of(paper)
    if key not in PAPER_LABELS:
        return None
    size, _dash, columns = key.partition("-")
    return {
        "size": size,
        "orientation": "landscape" if size in ("8K", "A3") else "portrait",
        "columns": int(columns) if columns.isdigit() else 1,
    }


def describe(profile):
    """把画像说成一段人话，给对话里的引导语用（也给人看）。

    只写**确实记住的**：没记住的节不占篇幅，免得模型以为"学校是空的"是个信息。
    """
    if not isinstance(profile, dict):
        profile, _path, _exists = load()
    lines = []
    school = profile.get("school") or {}
    bits = [school.get("stage"), school.get("grade"), school.get("subject")]
    who = "".join(str(item) for item in bits if item)
    if school.get("name") or who:
        text = str(school.get("name") or "")
        if who:
            text = (text + "，" if text else "") + who
        if school.get("schoolYear"):
            text += "，" + str(school["schoolYear"])
        lines.append("学校与学科：" + text)
    if school.get("teacher"):
        lines.append("出卷/署名用：" + str(school["teacher"]))
    paper = profile.get("paper") or {}
    paper_key_value = paper_key_of(paper)  # 两种写法都认：key 是内部名，size+方向+栏数是老师嘴里的叫法
    if paper_key_value:
        lines.append("试卷默认纸型：" + paper_label(paper_key_value))
    if paper.get("footer"):
        lines.append("页脚写法：" + str(paper["footer"]))
    if paper.get("note"):
        lines.append("卷面习惯：" + str(paper["note"]))
    plan = profile.get("plan") or {}
    if plan.get("layout"):
        lines.append("教案默认版式：" + PLAN_LAYOUTS.get(str(plan["layout"]), str(plan["layout"])))
    if plan.get("note"):
        lines.append("教案习惯：" + str(plan["note"]))
    deck = profile.get("deck") or {}
    if deck.get("template"):
        lines.append("课件模板：" + str(deck["template"]))
    # theme / titleSizePt / bodySizePt 是面板与引导语真正会记的键；只认 palette/fontSize
    # 时，老师明明存了「clean 主题、标题 40pt」，摘要里却一个字不显示（以为没记住）。
    piece = []
    if deck.get("theme"):
        piece.append("主题 " + str(deck["theme"]))
    if deck.get("palette"):
        piece.append("配色 " + str(deck["palette"]))
    if deck.get("titleSizePt"):
        piece.append("标题字号 " + str(deck["titleSizePt"]) + "pt")
    body_size = deck.get("bodySizePt") or deck.get("fontSize")
    if body_size:
        piece.append("正文字号 " + str(body_size) + "pt")
    if deck.get("slideSize"):
        piece.append("画布 " + str(deck["slideSize"]))
    if piece:
        lines.append("课件风格：" + " · ".join(piece))
    wechat = profile.get("wechat") or {}
    if wechat.get("tone") or wechat.get("length") or wechat.get("signature"):
        piece = [item for item in (wechat.get("tone"), wechat.get("length")) if item]
        text = "公众号语气：" + ("、".join(str(item) for item in piece) if piece else "默认")
        if wechat.get("signature"):
            text += "；落款 " + str(wechat["signature"])
        lines.append(text)
    docflow = profile.get("docflow") or {}
    if docflow.get("unit") or docflow.get("prefix") or docflow.get("tone"):
        piece = []
        if docflow.get("unit"):
            piece.append("本单位 " + str(docflow["unit"]))
        if docflow.get("prefix"):
            piece.append("字号前缀 " + str(docflow["prefix"]))
        if docflow.get("tone"):
            piece.append("批复语气 " + str(docflow["tone"]))
        lines.append("公文：" + " · ".join(piece))
    tasks = profile.get("tasks") or {}
    if tasks.get("minutes"):
        lines.append("常用考试时长：" + str(tasks["minutes"]) + " 分钟")
    return "\n".join("· " + line for line in lines)


def summary(profile=None):
    """给宿主与面板用的一份紧凑画像（不返回 history，省得刷屏）。"""
    if profile is None:
        profile, path, exists = load()
    else:
        path, exists = profile_path(), True
    slim = {name: dict(profile.get(name) or {}) for name in SECTIONS}
    return {
        "profile": slim,
        "path": path,
        "exists": exists,
        "text": describe(profile),
        "learnedCount": len(profile.get("learned") or []),
        "filled": [name for name in SECTIONS if profile.get(name)],
    }

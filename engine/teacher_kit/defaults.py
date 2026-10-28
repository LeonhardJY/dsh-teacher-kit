"""把画像里的默认值填进这次的参数里（老师说了的就不动）。

为什么放在引擎这一层
--------------------
对话模式里，参数是**模型按老师这句话**现攒的：它知道"这次七(3)班、90 分钟"，
但不知道"你们学校印 8 开两栏、页脚要页码"。这些默认值要么靠模型记住（不可靠），
要么靠宿主每次拼 prompt（会漂）。放在命令入口统一兜底，一处生效、可测。

规则就一条：**payload 里有的绝不动**。只有缺的、或只给了个空值时，才用画像补。
"""

from . import profile as profile_store

# 老师嘴里的叫法 → paper.resolve 认的 {size, orientation, columns}
PAPERS = {
    "8K-2": {"size": "8K", "orientation": "landscape", "columns": 2},
    "8K-1": {"size": "8K", "orientation": "landscape", "columns": 1},
    "A3-2": {"size": "A3", "orientation": "landscape", "columns": 2},
    "A3-1": {"size": "A3", "orientation": "landscape", "columns": 1},
    "16K-1": {"size": "16K", "orientation": "portrait", "columns": 1},
    "A4-1": {"size": "A4", "orientation": "portrait", "columns": 1},
}

# 别名：模型可能按"8开两栏""8K2""8k-2"这类写法给，都认
PAPER_ALIASES = {
    "8K2": "8K-2", "8开2": "8K-2", "8开两栏": "8K-2", "8开横向两栏": "8K-2",
    "8K1": "8K-1", "8开1": "8K-1", "8开一栏": "8K-1", "8开横向一栏": "8K-1",
    "A32": "A3-2", "A3两栏": "A3-2", "A3横向两栏": "A3-2",
    "A31": "A3-1", "A3一栏": "A3-1", "A3横向一栏": "A3-1",
    "16K1": "16K-1", "16开": "16K-1", "16开一栏": "16K-1",
    "A41": "A4-1", "A4": "A4-1", "A4一栏": "A4-1",
}

PLAN_ALIASES = {
    "表格式": "form", "表格": "form", "form": "form", "table": "form",
    "流程版": "flow", "流程": "flow", "flow": "flow",
}

PAPER_TASKS = ("exam:build", "exam:build-ab", "exam:build-teacher", "bank:build")
PLAN_TASKS = ("plan:build",)
DECK_TASKS = ("lesson:deck",)
HANDOUT_TASKS = ("handout:build",)
WECHAT_TASKS = ("wechat:render",)
DOCFLOW_TASKS = ("docflow:build",)


def paper_key(value):
    """把各种写法收敛成 PAPERS 的键；认不出来就返回 None（保持"没给"）。"""
    if not value:
        return None
    text = str(value).strip().replace(" ", "").replace("－", "-").replace("—", "-")
    upper = text.upper()
    if upper in PAPERS:
        return upper
    compact = upper.replace("-", "")
    if compact in PAPERS:
        return compact
    return PAPER_ALIASES.get(text) or PAPER_ALIASES.get(upper) or PAPER_ALIASES.get(compact)


def _text(value):
    if value in (None, ""):
        return ""
    # 画像里存成数组的习惯（classNames: ["七(1)班","七(2)班"]）不能把 Python 列表字面量印进页脚
    if isinstance(value, (list, tuple)):
        return "、".join(str(item).strip() for item in value if str(item).strip())
    return str(value).strip()


def _number(value):
    try:
        number = float(value)
    except (TypeError, ValueError):
        return None
    return int(number) if number == int(number) else number


def _fill_options(payload):
    options = payload.get("options")
    if not isinstance(options, dict):
        options = {}
        payload["options"] = options
    return options


def apply(payload, command, profile=None):
    """按画像补齐这次没说到的项，返回（可能被改过的）payload。

    只在**没有**值的时候补；画像是老师的"习惯"，不是强制。
    """
    if not isinstance(payload, dict):
        return {}
    if profile is None:
        profile, _path, _exists = profile_store.load()
    school = profile.get("school") or {}
    paper = profile.get("paper") or {}
    tasks = profile.get("tasks") or {}
    plan_cfg = profile.get("plan") or {}
    deck_cfg = profile.get("deck") or {}

    filled = []

    # ① 纸型：给了 baseTemplate 就按模板走（模板比画像权威：这次套的就是学校的卷子）
    if command in PAPER_TASKS and not payload.get("page") and not payload.get("baseTemplate"):
        page = profile_store.paper_page(payload.get("paper")) or profile_store.paper_page(paper)
        if page:
            payload["page"] = page
            filled.append("page=" + str(page.get("size")) + "x" + str(page.get("columns")))
    if command in PAPER_TASKS and payload.get("baseTemplate") and not payload.get("page"):
        # 套模板时把画像里的纸型也带上：模板读不出纸型时才有个像样的兜底
        page = profile_store.paper_page(paper)
        if page:
            payload["_profilePage"] = page

    # ② 页脚两格：这次没说就用画像里的（老师姓名 / 常用班级）
    if command in PAPER_TASKS:
        values = payload.get("values")
        if not isinstance(values, dict):
            values = {}
            payload["values"] = values
        if not _text(values.get("teacherName")):
            if _text(school.get("teacher")):
                values["teacherName"] = _text(school["teacher"])
                filled.append("teacherName")
        if not _text(values.get("className")) and _text(school.get("classNames")):
            values["className"] = _text(school["classNames"])
            filled.append("className")

        # 卷名：老师没说时按画像里的学科给个默认（「英语测试卷」）。
        # 以前缺省拿"第一个大题的标题"当卷名，卷首大字印成「一、单项选择」，老师得改了再印。
        if not _text(values.get("examTitle")) and not _text(payload.get("title")):
            subject = _text(school.get("subject"))
            if subject:
                values["examTitle"] = subject + "测试卷"
                filled.append("examTitle")

        # 页脚写法：画像里记了「使用班级：{{className}} 出卷老师：{{teacherName}}」这类模板就用它，
        # 以前只存在画像里从不注入，老师改了写法出卷还是内置那一行（实测回执说「页脚写法：…」但没用上）
        if not _text(payload.get("footer")) and _text(paper.get("footer")):
            payload["footer"] = _text(paper["footer"])
            filled.append("footer")

    # ③ 考试时长 / 答题卡
    if command in PAPER_TASKS + HANDOUT_TASKS:
        options = _fill_options(payload)
        if not options.get("minutes"):
            minutes = _number(tasks.get("minutes")) or _number(paper.get("minutes"))
            if minutes:
                options["minutes"] = minutes
                filled.append("minutes")
        if command in PAPER_TASKS and not options.get("answerCard"):
            want = tasks.get("answerCard")
            if isinstance(want, bool):
                options["answerCard"] = "card" if want else "none"
                filled.append("answerCard")

    # ④ 听写单：几栏 / 用哪种模式
    if command in HANDOUT_TASKS:
        if not payload.get("columns") and _number(tasks.get("columns")):
            payload["columns"] = _number(tasks["columns"])
            filled.append("columns")
        if not payload.get("mode") and _text(tasks.get("dictationMode")):
            payload["mode"] = _text(tasks["dictationMode"])
            filled.append("mode")

    # ⑤ 教案：表格式 / 流程版
    if command in PLAN_TASKS:
        if not payload.get("layout"):
            layout = PLAN_ALIASES.get(_text(plan_cfg.get("layout")))
            if layout:
                payload["layout"] = layout
                filled.append("layout")
        if not _text(payload.get("teacher")) and _text(school.get("teacher")):
            payload["teacher"] = _text(school["teacher"])
            filled.append("teacher")

    # ⑥ 课件：字号、画布比例（配色靠模板，不在这里造）
    if command in DECK_TASKS:
        options = _fill_options(payload)
        if not options.get("titleSizePt") and _number(deck_cfg.get("titleSizePt")):
            options["titleSizePt"] = _number(deck_cfg["titleSizePt"])
            filled.append("titleSizePt")
        if not options.get("bodySizePt") and _number(deck_cfg.get("fontSize")):
            options["bodySizePt"] = _number(deck_cfg["fontSize"])
            filled.append("bodySizePt")
        if not options.get("slideSize") and _text(deck_cfg.get("slideSize")):
            options["slideSize"] = _text(deck_cfg["slideSize"])
            filled.append("slideSize")
        if not _text(payload.get("template")) and _text(deck_cfg.get("template")):
            # 只在校内共享目录这种确实存在的路径上兜底：模板文件被挪走了就别硬套
            import os
            if os.path.isfile(_text(deck_cfg.get("template"))):
                payload["template"] = _text(deck_cfg["template"])
                filled.append("template")

    # ⑦ 公众号：落款
    if command in WECHAT_TASKS:
        wechat = profile.get("wechat") or {}
        if not _text(payload.get("signature")) and _text(wechat.get("signature")):
            payload["signature"] = _text(wechat["signature"])
            filled.append("signature")

    # ⑧ 公文：本单位全称、常打交道的科室
    if command in DOCFLOW_TASKS:
        docflow_cfg = profile.get("docflow") or {}
        if not _text(payload.get("unit")) and _text(docflow_cfg.get("unit")):
            payload["unit"] = _text(docflow_cfg["unit"])
            filled.append("unit")
        if not _text(payload.get("dept")) and _text(docflow_cfg.get("dept")):
            payload["dept"] = _text(docflow_cfg["dept"])
            filled.append("dept")

    if filled:
        payload["_profileFilled"] = filled
    return payload

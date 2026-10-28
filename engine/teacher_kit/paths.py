r"""文件放哪儿：老师不该为了存一份卷子去理解路径。

统一落在「文档/教师助手/<任务>/」下，文件名带上时间，重复出卷不会互相覆盖。
老师只需要在面板上点「打开文件夹」，不需要知道 C:\Users\... 是什么。
"""

import os
import re
import time

from .errors import EngineError

RE_BAD = re.compile(r'[\\/:*?"<>|\r\n\t]+')


def teacher_home():
    override = os.environ.get("TEACHER_KIT_HOME")
    if override:
        return override
    home = os.path.expanduser("~")
    documents = os.path.join(home, "Documents")
    base = documents if os.path.isdir(documents) else home
    return os.path.join(base, "教师助手")


def safe_name(text, fallback="未命名"):
    name = RE_BAD.sub("_", str(text or "")).strip(" ._")
    name = re.sub(r"\s+", " ", name)
    return (name or fallback)[:60]


def guard_output_path(path):
    """显式 outputPath 也要过一遍：Windows 上名字里带半角冒号会写成 NTFS 数据流，
    文件夹里只剩一个 0 字节的空壳，36KB 的成品谁都看不见（实测）。"""
    raw = str(path or "")
    if "\x00" in raw:
        # 路径里混进 NUL（JSON 里的 \u0000）：一路穿到 zipfile 会报英文 embedded null character
        raise EngineError(
            "bad-output-path",
            "文件名里有个看不见的控制字符（\u0000）：" + os.path.basename(raw.replace("\\", "/")),
            "把文件名重新打一遍；从别处复制来的名字可能带着看不见的字符")
    base = os.path.basename(raw.replace("\\", "/"))
    if not base:
        return
    bad = RE_BAD.search(base)
    if bad:
        raise EngineError(
            "bad-output-path",
            "文件名里有不能用的符号「" + bad.group(0) + "」：" + base,
            "去掉 \\ / : * ? \" < > | 这些符号再存（用「-」或全角「：」都行）")


def default_output(kind, title, extension):
    """kind = 出卷 / 课件 / 文件加工

    精确到秒，重名再补 -2/-3：一分钟内连着出两份卷子不许互相覆盖
    （老师手工微调过的那份输不起）。
    """
    folder = os.path.join(teacher_home(), safe_name(kind, "教师助手"))
    os.makedirs(folder, exist_ok=True)
    # 时间戳带毫秒：同一秒里连点两次「出卷」（两个引擎进程）不会选中同一个文件名。
    # 毫秒还撞上（极少）就靠下面的 -2/-3 兜底。
    stamp = time.strftime("%Y%m%d-%H%M%S") + "-%03d" % (time.time_ns() // 1_000_000 % 1000)
    base = os.path.join(folder, safe_name(title) + "_" + stamp)
    candidate = base + extension
    suffix = 2
    while os.path.exists(candidate):
        candidate = base + "-" + str(suffix) + extension
        suffix += 1
    return candidate
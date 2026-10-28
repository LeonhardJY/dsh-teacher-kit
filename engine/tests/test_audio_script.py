"""听力录音稿音频（audio_script）的回归测试。

依赖 Windows 自带语音包：没语音的机器打印说明并跳过真合成，不让整套测试变红。
"""

import os
import sys
import tempfile
import wave

HERE = os.path.dirname(os.path.abspath(__file__))
ENGINE = os.path.dirname(HERE)
sys.path.insert(0, ENGINE)

from teacher_kit.audio_script import _segments, build_script_audio  # noqa: E402
from teacher_kit.dictation_audio import list_voices  # noqa: E402
from teacher_kit.errors import EngineError  # noqa: E402

PASSED = 0
FAILED = 0
SKIPPED = []


def check(cond, label, detail=""):
    global PASSED, FAILED
    if cond:
        PASSED += 1
        print("  ✓ " + label)
    else:
        FAILED += 1
        print("  ✗ " + label + ("  ← " + str(detail) if detail else ""))


def expect_error(code, fn, label):
    try:
        fn()
    except EngineError as exc:
        check(exc.code == code, label, "实际 code=" + str(exc.code) + " message=" + str(exc.message))
        check(bool(exc.message), label + "（有中文说明）", exc.message)
        return
    check(False, label, "没有抛 EngineError")


def wav_info(path):
    with wave.open(path, "rb") as fh:
        return {
            "channels": fh.getnchannels(),
            "width": fh.getsampwidth(),
            "rate": fh.getframerate(),
            "seconds": round(fh.getnframes() / float(fh.getframerate()), 1),
        }


def main():
    tmp = tempfile.mkdtemp(prefix="tk-script-")
    os.environ["TEACHER_KIT_HOME"] = tmp

    print("[1] 中英混排切成两段语音")
    parts = _segments("Hello 你好 world")
    check([p[0] for p in parts] == ["en", "zh", "en"], "英文/中文/英文各成一段", parts)
    check(parts[1][1] == "你好", "中文段内容正确", parts)
    check(_segments("   ") == [], "全是空白没有段", _segments("   "))

    print("[2] 参数校验（这些错要在花时间合成之前就拦住）")
    expect_error("text-required", lambda: build_script_audio({"text": "   "}), "空稿子报 text-required")
    expect_error("bad-pause", lambda: build_script_audio({"text": "Hello", "pauseSec": 99}), "停顿超范围报 bad-pause")
    expect_error("bad-gap", lambda: build_script_audio({"text": "Hello", "gapSec": 30}), "句间停顿超范围报 bad-gap")
    expect_error("bad-lead", lambda: build_script_audio({"text": "Hello", "leadSec": 60}), "开头留白超范围报 bad-lead")
    expect_error("bad-rate", lambda: build_script_audio({"text": "Hello", "rate": 20}), "语速超范围报 bad-rate")
    expect_error("too-long",
                 lambda: build_script_audio({"text": "\n".join(["Hello there, this is a long line."] * 400)}),
                 "稿子太长报 too-long")

    print("[3] 真合成：三行（含中文）→ wav 头与时长")
    try:
        voices = list_voices()
    except EngineError as exc:
        voices = []
        SKIPPED.append("list_voices: " + str(exc.message))
    if not voices:
        print("  这台机器没有可用的语音包，跳过真合成（不影响其它测试）")
        print()
        print("全部通过 ✅（跳过 " + str(len(SKIPPED)) + " 组）" if FAILED == 0 else "失败 " + str(FAILED) + " 项")
        return 0 if FAILED == 0 else 1

    script = "1. What does the boy want to buy?\n他想要买一个书包。\n2. Where are they talking?"
    first = build_script_audio({"text": script, "number": True, "title": "听力试做", "pauseSec": 3})
    info = wav_info(first["outputPath"])
    check(info["channels"] == 1, "单声道", info)
    check(info["width"] == 2, "16 bit", info)
    check(info["rate"] == 22050, "22.05kHz", info)
    check(abs(info["seconds"] - first["seconds"]) < 1.0, "时长和 wav 头对得上", (info["seconds"], first["seconds"]))
    check(info["seconds"] > 8, "三行念下来超过 8 秒", info["seconds"])
    check(first["lines"] == 3, "按 3 行念", first["lines"])
    check(first["bytes"] > 10000, "文件不是空的", first["bytes"])
    check(os.path.exists(first["outputPath"]), "音频落在磁盘上", first["outputPath"])
    check(first["voices"]["en"] or first["voices"]["zh"], "报了用的是哪个语音", first["voices"])

    print("[4] 停顿调长 → 音频更长")
    short = build_script_audio({"text": script, "pauseSec": 1, "title": "短停"})
    long_gap = build_script_audio({"text": script, "pauseSec": 9, "title": "长停"})
    check(long_gap["seconds"] > short["seconds"] + 10, "每句停 9 秒比停 1 秒长 10 秒以上",
          (short["seconds"], long_gap["seconds"]))

    print("[5] 念题号会更长（多念了 Number one.）")
    plain = build_script_audio({"text": script, "number": False, "title": "不念题号"})
    check(first["seconds"] > plain["seconds"], "念题号比不念长", (first["seconds"], plain["seconds"]))

    print("[6] 整段念两遍")
    twice = build_script_audio({"text": script, "passes": 2, "title": "念两遍", "number": False})
    once = build_script_audio({"text": script, "passes": 1, "title": "念一遍", "number": False})
    check(twice["passes"] == 2, "回传念了几遍", twice.get("passes"))
    check(twice["seconds"] > once["seconds"] * 1.6, "念两遍比一遍长一倍左右",
          (once["seconds"], twice["seconds"]))
    expect_error("bad-passes",
                 lambda: build_script_audio({"text": script, "passes": 5, "title": "遍数越界"}),
                 "念 5 遍要报错")

    print("[7] 同名文件不覆盖")
    expect_error("output-exists",
                 lambda: build_script_audio({"text": script, "outputPath": first["outputPath"]}),
                 "同名音频已存在要报错，不能悄悄盖掉")
    again = build_script_audio({"text": script, "outputPath": first["outputPath"], "overwrite": True})
    check(again["seconds"] > 0, "明确 overwrite=True 才覆盖", again["seconds"])

    print()
    if FAILED:
        print("失败 " + str(FAILED) + " 项")
        return 1
    print("全部通过 ✅" + ("（跳过 " + str(len(SKIPPED)) + " 组）" if SKIPPED else ""))
    return 0


if __name__ == "__main__":
    sys.exit(main())

r"""听力录音稿 → 音频：老师把听力原文粘进来，用 Windows 自带的语音念成 wav。

为什么不直接给 mp3：本机没有 mp3 编码器（ffmpeg/lame 都没装），wav 能直接播、能进 PPT。
一段听力 3 分钟约 8MB，讲课用够了；要 mp3 再用格式工厂转一次。

和听写音频的区别：听写是"念一个词 → 停几秒 → 再念一个词"，听力是"按稿子连读，
句子之间留可控的停顿"。所以这里按行念，行内中英混排会自动切到对应语言的语音。
"""

import json
import os
import re
import tempfile
import wave

from .dictation_audio import (
    CHANNELS,
    PS_SCRIPT,
    SAMPLE_RATE,
    SAMPLE_WIDTH,
    _pitch_voice,
    _powershell,
    _run_ps,
    _silence,
    _wav_bytes,
    list_voices,
)
from .errors import EngineError, require
from .handout import CJK_RANGE
from .paths import default_output

DEFAULT_GAP = 1.2          # 一句念完到下一句之间：连读感，不打断语流
DEFAULT_PAUSE = 4.0        # 一"题"念完到下一题之间：听力考试留的作答时间
DEFAULT_LEAD = 2.0
DEFAULT_TAIL = 2.0
DEFAULT_RATE = 0
MAX_LINES = 400
MAX_TOTAL_SECONDS = 60 * 40
READ_CHARS_PER_SECOND = 11.0   # 中文约 5 字/秒、英文约 15 字母/秒，取中间估时长

TEXT_HINT = "把听力原文整段粘进来（一行一句，题与题之间空一行也可以）"

RE_CJK = re.compile("[" + CJK_RANGE + "]")
# 一行里中英混排时按"连续同文字体系"切段，各段交给对应的语音
RE_RUN = re.compile("([" + CJK_RANGE + r"\u3000-\u303f\uff00-\uffef]+|[^" + CJK_RANGE + r"\u3000-\u303f\uff00-\uffef]+)")
RE_NUMBERING = re.compile(r"^\s*(?:[(（]\s*(\d{1,3})\s*[)）]|(\d{1,3})\s*[.、)）．])\s*")
RE_SPOKEN = re.compile(r"[^\s\u3000]")

EN_NUMBERS = ("one", "two", "three", "four", "five", "six", "seven", "eight", "nine", "ten",
              "eleven", "twelve", "thirteen", "fourteen", "fifteen", "sixteen", "seventeen",
              "eighteen", "nineteen", "twenty")


def _spoken_length(text):
    """能念出来的字符数（空白不算）。"""
    return len(RE_SPOKEN.findall(text or ""))


def _segments(line):
    """一行 → [(语言, 文本)]：中文段走中文语音，英文段走英文语音。"""
    parts = []
    for chunk in RE_RUN.findall(line):
        text = chunk.strip(" \u3000\t")
        if not text or not _spoken_length(text):
            continue
        parts.append(("zh" if RE_CJK.search(text) else "en", text))
    return parts


def _en_number_word(value):
    if 1 <= value <= len(EN_NUMBERS):
        return EN_NUMBERS[value - 1]
    return str(value)


def build_script_audio(payload):
    """payload = {text, title?, outputPath?, overwrite?, number?, gapSec?, pauseSec?, leadSec?, tailSec?, rate?, passes?, enVoice?, zhVoice?}

    text 一行念一句；number=True 时把行首的 1. / 1、 剥掉，并先用英文念"Number one."；
    passes=2 表示整段念两遍（中间停 pauseSec 秒，听力考试常见）。
    返回 {outputPath, seconds, bytes, lines, chars, passes, voices, warnings}
    """
    payload = payload or {}
    text = payload.get("text") or ""
    require(bool(text.strip()), "text-required", "没有收到听力录音稿", TEXT_HINT)

    raw_lines = [line.strip() for line in text.replace("\r\n", "\n").replace("\r", "\n").split("\n")]
    lines = [line for line in raw_lines if line]
    require(bool(lines), "no-items", "这段文字里一句都没认出来", TEXT_HINT)

    warnings = []
    if len(lines) > MAX_LINES:
        warnings.append({"code": "too-many-lines",
                         "message": "一次最多念 " + str(MAX_LINES) + " 行，后面的没有念"})
        lines = lines[:MAX_LINES]

    gap = float(payload.get("gapSec", DEFAULT_GAP))
    if gap < 0 or gap > 10:
        raise EngineError("bad-gap", "句子之间的停顿只能是 0—10 秒：" + str(payload.get("gapSec")),
                          "连读念稿子用 1—2 秒")
    pause = float(payload.get("pauseSec", DEFAULT_PAUSE))
    if pause < 0 or pause > 60:
        raise EngineError("bad-pause", "每句后面的停顿只能是 0—60 秒：" + str(payload.get("pauseSec")),
                          "听力考试留作答时间用 3—8 秒")
    lead = float(payload.get("leadSec", DEFAULT_LEAD))
    if lead < 0 or lead > 30:
        raise EngineError("bad-lead", "开头留白只能是 0—30 秒：" + str(payload.get("leadSec")), "默认 2 秒")
    tail = float(payload.get("tailSec", DEFAULT_TAIL))
    if tail < 0 or tail > 30:
        raise EngineError("bad-tail", "结尾留白只能是 0—30 秒：" + str(payload.get("tailSec")), "默认 2 秒")
    passes = int(payload.get("passes", 1) or 1)
    if passes < 1 or passes > 3:
        raise EngineError("bad-passes", "整段念几遍只能是 1—3：" + str(payload.get("passes")),
                          "听力考试通常整段放两遍，中间留作答时间")
    rate = int(payload.get("rate", DEFAULT_RATE) or 0)
    if rate < -10 or rate > 10:
        raise EngineError("bad-rate", "语速只能是 -10（最慢）到 10（最快）：" + str(payload.get("rate")), "慢一点给 -2")
    numbered = bool(payload.get("number", False))

    total_chars = sum(_spoken_length(line) for line in lines)
    estimate = lead + tail + passes * (total_chars / READ_CHARS_PER_SECOND + len(lines) * (pause + gap))
    if estimate > MAX_TOTAL_SECONDS:
        raise EngineError("too-long",
                          "这段稿子念下来要约 " + str(round(estimate / 60)) + " 分钟，太长了",
                          "分成两段分别生成，或者把语速调快一点")

    voices = list_voices()
    en_voice = payload.get("enVoice") or _pitch_voice(voices, "en")
    zh_voice = payload.get("zhVoice") or _pitch_voice(voices, "zh")
    if not en_voice and not zh_voice:
        raise EngineError("tts-no-voice", "这台电脑上没有装任何语音包，念不出声音",
                          "设置 → 时间和语言 → 语音 → 添加语音（英语 + 中文各一个）")
    if not en_voice:
        warnings.append({"code": "no-en-voice", "message": "没找到英文语音，英文句子会用中文语音念（发音不准）"})

    title = str(payload.get("title") or "").strip() or "听力录音"
    # 名字里已经带「音频 / 录音」的就别再叠一层，免得出现「X 听力 听力音频.wav」
    name = title if ("音频" in title or "录音" in title) else title + " 听力音频"
    output = payload.get("outputPath") or default_output("听写", name, ".wav")
    if os.path.exists(output) and not payload.get("overwrite"):
        raise EngineError("output-exists", "同名音频已经存在了：" + os.path.basename(output),
                          "换个名字，或者把已有的删掉再来一次")

    exe = _powershell()
    jobs = []
    with tempfile.TemporaryDirectory(prefix="tk-script-") as tmp:
        index = 0
        for ordinal, line in enumerate(lines, start=1):
            head = ""
            body = line
            match = RE_NUMBERING.match(line)
            if match:
                body = line[match.end():].strip()
                if numbered:
                    number_text = (match.group(1) or match.group(2) or "").strip()
                    if number_text:
                        head = "Number " + _en_number_word(int(number_text))

            pieces = []
            if head:
                pieces.append(("en", head))
            pieces.extend(_segments(body))

            for position, (language, content) in enumerate(pieces):
                voice = (zh_voice if language == "zh" else en_voice) or en_voice or zh_voice
                index += 1
                jobs.append({
                    "text": content,
                    "file": os.path.join(tmp, "part-%04d.wav" % index),
                    "voice": voice,
                    "line": ordinal,
                    "last": position == len(pieces) - 1,
                })

        groups = {}
        for job in jobs:
            groups.setdefault(job["voice"], []).append({"text": job["text"], "file": job["file"]})
        if not groups:
            raise EngineError("no-items", "这段文字里没有能念的内容", TEXT_HINT)

        script = os.path.join(tmp, "synth.ps1")
        with open(script, "w", encoding="utf-8") as handle:
            handle.write(PS_SCRIPT)
        for voice, entries in groups.items():
            data_file = os.path.join(tmp, "lines-%d.json" % len(os.listdir(tmp)))
            with open(data_file, "w", encoding="utf-8") as handle:
                json.dump(entries, handle, ensure_ascii=False)
            _run_ps(exe, ["-File", script, "-InputFile", data_file, "-Voice", voice, "-Rate", str(rate)])

        params = None
        body = []
        for job in jobs:
            if not os.path.exists(job["file"]) or os.path.getsize(job["file"]) < 44:
                warnings.append({"code": "part-missing",
                                 "message": "有一句没念出来（第 " + str(job["line"]) + " 行），已跳过"})
                continue
            if params is None:
                params = _wav_bytes(job["file"])[0]
            body.append(_wav_bytes(job["file"])[1])
            body.append(_silence(pause if job["last"] else gap))
        if params is None:
            raise EngineError("tts-failed", "一段音频都没合成出来",
                              "先确认系统里有语音包（设置 → 时间和语言 → 语音）")
        # passes=2：整段念完停 pause 秒再来一遍（听力考试最常见的用法）
        chunks = [_silence(lead)]
        for round_index in range(passes):
            chunks.extend(body)
            if round_index < passes - 1:
                chunks.append(_silence(pause))
        chunks.append(_silence(tail))

        directory = os.path.dirname(os.path.abspath(output))
        if directory and not os.path.isdir(directory):
            os.makedirs(directory, exist_ok=True)
        with wave.open(output, "wb") as out:
            out.setnchannels(params[0])
            out.setsampwidth(params[1])
            out.setframerate(params[2])
            for chunk in chunks:
                out.writeframes(chunk)

    seconds = round(sum(len(chunk) for chunk in chunks) / float(params[2] * params[1] * params[0]), 1)
    return {
        "outputPath": output,
        "seconds": seconds,
        "minutes": round(seconds / 60.0, 1),
        "bytes": os.path.getsize(output),
        "lines": len(lines),
        "chars": total_chars,
        "passes": passes,
        "voices": {"en": en_voice, "zh": zh_voice},
        "warnings": warnings,
    }

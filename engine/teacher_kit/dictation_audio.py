"""听写音频：把单词表念成一段能直接放的 wav。

为什么用 Windows 自带的语音（System.Speech / SAPI）：
- 老师的电脑上一定有，不用装任何东西、不联网、不花钱——这正是这个插件的底线
  （引擎只用 DSH 自带的 Python 与 LibreOffice）。
- 代价是只能出 **wav**：SAPI 不带你机器上的 mp3 编码器（要有 LAME/ffmpeg 才行）。
  听写音频就几十秒、只在教室里放一次，wav 完全够用；22.05kHz 16bit 单声道
  （约 2.6MB/分钟），U 盘拷进班班通也不费劲。

为什么一次 PowerShell 调用念完整张表：
逐条起进程要 0.3s/条，40 个词就是十几秒；把"念哪些句子、存哪个文件"写成一份
JSON 交给 PowerShell，一个进程内循环，整张表 1—3 秒。

排版侧的约定（阅读顺序 = 播放顺序）：
每一条念 repeat 遍 → 停 pauseSec 秒（学生写）→ 下一条。开头留 leadSec 秒，
让老师点完播放还来得及走到讲台。题号用「第 N 题」（中文）/「Number N」（英文），
和听写单上印的序号对得上。
"""

import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import wave

from .errors import EngineError, require
from .handout import (MODE_HINT, TEXT_HINT, _norm_mode, _to_int, parse_items,
                     read_word_list)
from .paths import default_output

MAX_ITEMS = 120
SAMPLE_RATE = 22050
SAMPLE_WIDTH = 2          # 16 bit
CHANNELS = 1
DEFAULT_REPEAT = 2        # 念两遍：第一遍没听清还有第二遍，是听写的常规做法
DEFAULT_PAUSE = 6         # 6 秒够七年级学生写完一个单词（词组长一点就调 8）
DEFAULT_LEAD = 3
DEFAULT_RATE = 0          # -10（慢）— 10（快）；0 是正常语速
MAX_TOTAL_SECONDS = 60 * 30   # 一段音频超过 30 分钟就不像听写了，多半是词表选错

CJK_RANGE = "\u3400-\u4dbf\u4e00-\u9fff\uf900-\ufaff"
RE_CJK = re.compile("[" + CJK_RANGE + "]")

# 1—120 的英文读法：SAPI 念 "1" 会读 one，但 "第 1 题" 混着念更自然，也避免数字被念成
# "one hundred and twenty" 时和条目粘在一起。只覆盖到 120（一次一份的上限）。
_EN_ONES = ("zero one two three four five six seven eight nine ten eleven twelve thirteen "
            "fourteen fifteen sixteen seventeen eighteen nineteen").split()
_EN_TENS = ("", "", "twenty", "thirty", "forty", "fifty", "sixty", "seventy", "eighty", "ninety")


def _en_number(value):
    if value < 20:
        return _EN_ONES[value]
    if value < 100:
        tens, ones = divmod(value, 10)
        return _EN_TENS[tens] + ("-" + _EN_ONES[ones] if ones else "")
    hundreds, rest = divmod(value, 100)
    head = _EN_ONES[hundreds] + " hundred"
    if not rest:
        return head
    return head + " and " + _en_number(rest)


def _pitch_voice(voices, language):
    """在已安装的语音里挑一条够用的。返回 None 表示这台机器没有对应语言的语音。"""
    wanted = ("zh", "cn", "chinese", "huihui", "yaoyao", "kangkang") if language == "zh" \
        else ("en", "english", "zira", "david", "mark", "hazel")
    for voice in voices:
        name = (voice.get("name") or "").lower()
        culture = (voice.get("culture") or "").lower()
        if any(token in name or token in culture for token in wanted):
            return voice.get("name")
    return None


PS_LIST_VOICES = r"""
Add-Type -AssemblyName System.Speech
$s = New-Object System.Speech.Synthesis.SpeechSynthesizer
$s.GetInstalledVoices() | ForEach-Object {
  if ($_.Enabled) {
    $i = $_.VoiceInfo
    [pscustomobject]@{ name = $i.Name; culture = $i.Culture.Name } | ConvertTo-Json -Compress
  }
}
$s.Dispose()
"""

PS_SCRIPT = r"""
param([string]$InputFile, [string]$Voice, [int]$Rate)
Add-Type -AssemblyName System.Speech
$s = New-Object System.Speech.Synthesis.SpeechSynthesizer
$fmt = New-Object System.Speech.AudioFormat.SpeechAudioFormatInfo(22050, [System.Speech.AudioFormat.AudioBitsPerSample]::Sixteen, [System.Speech.AudioFormat.AudioChannel]::Mono)
if ($Voice) { $s.SelectVoice($Voice) }
$s.Rate = $Rate
$raw = [System.IO.File]::ReadAllText($InputFile, [System.Text.Encoding]::UTF8)
$items = $raw | ConvertFrom-Json
foreach ($it in $items) {
  $s.SetOutputToWaveFile($it.file, $fmt)
  $s.Speak([string]$it.text)
}
$s.SetOutputToNull()
$s.Dispose()
"""


def _powershell():
    if sys.platform != "win32":
        raise EngineError(
            "tts-unavailable",
            "这个功能用的是 Windows 自带的语音，当前系统上用不了",
            "在 Windows 上运行 DSH 就能生成听写音频；其它系统可以先用听写单让学生自己读",
        )
    exe = shutil.which("powershell") or shutil.which("pwsh") or r"C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"
    if not exe or not os.path.exists(exe):
        raise EngineError("tts-unavailable", "找不到 PowerShell，念不出音频", "换一台 Windows 电脑试试，或者先只出听写单")
    return exe


def _run_ps(exe, args):
    try:
        done = subprocess.run([exe, "-NoProfile", "-ExecutionPolicy", "Bypass"] + args,
                              capture_output=True, timeout=180)
    except subprocess.TimeoutExpired:
        raise EngineError("tts-timeout", "语音合成超时了（这张表可能太长）",
                          "把词表分成两份，或者把条目数调到 60 以内")
    if done.returncode != 0:
        detail = (done.stderr or b"").decode("utf-8", "replace").strip().splitlines()
        raise EngineError("tts-failed", "语音合成没跑通：" + (detail[-1] if detail else "未知错误"),
                          "先手动试一次：设置 → 时间和语言 → 语音，看系统里有没有英文语音包")
    return (done.stdout or b"").decode("utf-8", "replace")


def list_voices():
    """这台机器上能用的语音（名字 + 语言），面板/自检用。"""
    exe = _powershell()
    with tempfile.TemporaryDirectory(prefix="tk-voices-") as tmp:
        script = os.path.join(tmp, "voices.ps1")
        with open(script, "w", encoding="utf-8") as fh:
            fh.write(PS_LIST_VOICES)
        output = _run_ps(exe, ["-File", script])
    voices = []
    for line in output.splitlines():
        line = line.strip()
        if not line.startswith("{"):
            continue
        try:
            item = json.loads(line)
        except ValueError:
            continue
        if item.get("name"):
            voices.append({"name": item["name"], "culture": item.get("culture") or ""})
    return voices


def _silence(seconds):
    return b"\x00" * int(max(0.0, seconds) * SAMPLE_RATE * SAMPLE_WIDTH * CHANNELS)


def _wav_bytes(path):
    with wave.open(path, "rb") as fh:
        params = (fh.getnchannels(), fh.getsampwidth(), fh.getframerate())
        return params, fh.readframes(fh.getnframes())


def _line_for(item, repeat, number, rate_extra=0.0):
    """一条听写的念法：题号 + 词/短语 × repeat。"""
    prompt = str(item.get("prompt") or "").strip()
    if not prompt:
        return []
    lines = []
    if number:
        if RE_CJK.search(prompt):
            lines.append("第 " + str(item.get("index")) + " 题")
        else:
            lines.append("Number " + _en_number(int(item.get("index") or 1)))
    lines.extend([prompt] * max(1, repeat))
    return lines


def build_dictation_audio(payload):
    """payload = {text, source?, mode?, title?, outputPath?, overwrite?, repeat?, pauseSec?,
                  leadSec?, number?, rate?, enVoice?, zhVoice?, count?}

    text 为空、给了 source（词表文件）时先读文件。

    返回 {outputPath, count, seconds, bytes, voices, mode, warnings}
    """
    payload = payload or {}
    text = payload.get("text")
    if not (text or "").strip() and payload.get("source"):
        text = read_word_list(payload.get("source"))
    require(bool((text or "").strip()), "text-required", "没有收到听写内容", TEXT_HINT)

    mode = _norm_mode(payload.get("mode"))
    repeat = int(payload.get("repeat", DEFAULT_REPEAT) or DEFAULT_REPEAT)
    if repeat < 1 or repeat > 4:
        raise EngineError("bad-repeat", "念几遍只能是 1—4：" + str(payload.get("repeat")), "听写通常念 2 遍")
    pause = float(payload.get("pauseSec", DEFAULT_PAUSE) or 0)
    if pause < 1 or pause > 30:
        raise EngineError("bad-pause", "每条后面的停顿只能是 1—30 秒：" + str(payload.get("pauseSec")),
                          "七年级写一个单词 5—6 秒够，词组长一点给 8—10 秒")
    lead = float(payload.get("leadSec", DEFAULT_LEAD) or 0)
    if lead < 0 or lead > 30:
        raise EngineError("bad-lead", "开头留白只能是 0—30 秒：" + str(payload.get("leadSec")), "默认 3 秒")
    rate = int(payload.get("rate", DEFAULT_RATE) or 0)
    if rate < -10 or rate > 10:
        raise EngineError("bad-rate", "语速只能是 -10（最慢）到 10（最快）：" + str(payload.get("rate")), "慢一点给 -2")
    number = bool(payload.get("number", True))

    parsed = parse_items(text, mode)
    items = parsed["items"]
    warnings = list(parsed["warnings"])
    require(bool(items), "no-items", "这段文字里一条听写都没认出来",
            "检查是不是整段连在一起了；" + TEXT_HINT)
    if len(items) > MAX_ITEMS:
        warnings.append({"code": "too-many",
                         "message": "一次最多念 " + str(MAX_ITEMS) + " 条，后面的没有念"})
        items = items[:MAX_ITEMS]

    # count 是老师说的"只要前 N 条"：听写单和音频必须听同一句话——
    # 以前音频不看 count，单子 10 条、音频 16 条，对不上（评测 P0-2）。
    # 判据、区间与提示文案照抄 handout.build_dictation，两边不许长歪。
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

    total_guess = lead + len(items) * (pause + 3 * repeat)
    if total_guess > MAX_TOTAL_SECONDS:
        raise EngineError("too-long", "这张表念下来要 " + str(round(total_guess / 60)) + " 分钟，太长了",
                          "分成两份，或者把每条的停顿调短一点")

    voices = list_voices()
    en_voice = payload.get("enVoice") or _pitch_voice(voices, "en")
    zh_voice = payload.get("zhVoice") or _pitch_voice(voices, "zh")
    if not en_voice and not zh_voice:
        raise EngineError("tts-no-voice", "这台电脑上没有装任何语音包，念不出声音",
                          "设置 → 时间和语言 → 语音 → 添加语音（英语 + 中文各一个）")
    if not en_voice:
        warnings.append({"code": "no-en-voice", "message": "没找到英文语音，英文条目会用中文语音念（发音不准）"})
    if not zh_voice:
        warnings.append({"code": "no-zh-voice", "message": "没找到中文语音，题号会用英文语音念"})

    title = str(payload.get("title") or "").strip() or "英语听写"
    output = payload.get("outputPath") or default_output("听写", str(title) + " 音频", ".wav")
    if os.path.exists(output) and not payload.get("overwrite"):
        raise EngineError("output-exists", "同名音频已经存在了：" + os.path.basename(output),
                          "换个名字，或者把已有的删掉再来一次")

    exe = _powershell()
    jobs = []          # [{text, file, voice}]，按播放顺序；voice 为空表示这段是静音
    with tempfile.TemporaryDirectory(prefix="tk-tts-") as tmp:
        index = 0
        for item in items:
            lines = _line_for(item, repeat, number)
            for position, line in enumerate(lines):
                voice = zh_voice if RE_CJK.search(line) else en_voice
                if not voice:
                    voice = en_voice or zh_voice
                index += 1
                jobs.append({
                    "text": line,
                    "file": os.path.join(tmp, "part-%04d.wav" % index),
                    "voice": voice,
                    "item": item.get("index"),
                    "last": position == len(lines) - 1,
                })

        # 按语音分组：SelectVoice 换一次要几十毫秒，同一条语音的句子连着念最省时间
        groups = {}
        for job in jobs:
            groups.setdefault(job["voice"], []).append({"text": job["text"], "file": job["file"]})
        script = os.path.join(tmp, "synth.ps1")
        with open(script, "w", encoding="utf-8") as fh:
            fh.write(PS_SCRIPT)
        for voice, entries in groups.items():
            data_file = os.path.join(tmp, "lines-%d.json" % len(os.listdir(tmp)))
            with open(data_file, "w", encoding="utf-8") as fh:
                json.dump(entries, fh, ensure_ascii=False)
            _run_ps(exe, ["-File", script, "-InputFile", data_file, "-Voice", voice, "-Rate", str(rate)])

        # 拼接：开头留白 →（题号/词）段 → 停顿 → …… → 结尾留 2 秒
        # 同一条目内（题号后、念第几遍之间）用短停 0.8 秒，条目结束才用老师设的长停顿
        params = None
        chunks = [_silence(lead)]
        gap = _silence(0.8)
        for job in jobs:
            if not os.path.exists(job["file"]) or os.path.getsize(job["file"]) < 44:
                warnings.append({"code": "part-missing",
                                 "message": "有一段没念出来（" + os.path.basename(job["file"]) + "），已跳过"})
                continue
            if params is None:
                params = _wav_bytes(job["file"])[0]
            chunks.append(_wav_bytes(job["file"])[1])
            chunks.append(_silence(pause) if job["last"] else gap)
        if params is None:
            raise EngineError("tts-failed", "一段音频都没合成出来",
                              "先确认系统里有语音包（设置 → 时间和语言 → 语音）")
        chunks.append(_silence(2.0))

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
        "count": len(items),
        "seconds": seconds,
        "bytes": os.path.getsize(output),
        "voices": {"en": en_voice, "zh": zh_voice},
        "mode": mode,
        "warnings": warnings,
    }

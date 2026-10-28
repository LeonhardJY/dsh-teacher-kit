#!/usr/bin/env python
"""teacher-kit 文档引擎 CLI —— JSON 进，JSON 出。

用法：
  python cli.py <command> --args '<json>'
  python cli.py <command> --args-file <path>
  echo '<json>' | python cli.py <command> --args -

约定：
  * stdout 只输出**一行** JSON，宿主直接 JSON.parse。
  * 错误不抛栈到 stdout；栈只进 stderr（--debug 时才打）。
  * 所有给人看的文字是中文。
"""

import json
import os
import sys
import traceback

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from teacher_kit import __version__                      # noqa: E402
from teacher_kit.errors import EngineError               # noqa: E402
from teacher_kit import commands                         # noqa: E402
from teacher_kit import defaults                          # noqa: E402


def _load_json(text, label):
    """参数 JSON 坏了要说人话。

    以前截断的 JSON 直接抛 json 的英文异常进 internal 分支：
    「Expecting ',' delimiter: line 1 column 24」——老师/宿主都看不懂。
    """
    try:
        return json.loads(text)
    except ValueError as error:
        raise EngineError(
            "bad-json",
            label + "不是合法的 JSON（" + str(error)[:80] + "）",
            "多半是引号/逗号没配对，或者内容太长从中间断了；把这段参数重新发一次")


def _parse_args(argv):
    if len(argv) < 2:
        raise EngineError("missing-command", "没有给出命令", "可用命令：" + ", ".join(sorted(commands.TABLE)))
    command = argv[1]
    payload = {}
    index = 2
    while index < len(argv):
        token = argv[index]
        if token in ("--args", "--args-file"):
            if index + 1 >= len(argv):
                raise EngineError("missing-argument", token + " 后面没有值")
            value = argv[index + 1]
            if token == "--args-file":
                if not os.path.isfile(value):
                    raise EngineError(
                        "args-file-missing",
                        "找不到这个参数文件：" + os.path.basename(value),
                        "确认文件还在原位置；也可以直接 --args '{...}' 把参数写在命令里")
                with open(value, "r", encoding="utf-8") as handle:
                    payload = _load_json(handle.read(), "参数文件里的内容")
            elif value == "-":
                raw = sys.stdin.read()
                payload = _load_json(raw, "stdin 里的参数") if raw.strip() else {}
            else:
                payload = _load_json(value, "命令行里的参数")
            index += 2
            continue
        if token == "--debug":
            index += 1  # 只是打开栈输出，不是命令参数；别让它把整条命令打回 unknown-argument
            continue
        raise EngineError("unknown-argument", "不认识的参数：" + token)
    if not isinstance(payload, dict):
        # 参数不是对象时，以前要在 handler 里炸成英文 TypeError（32 条命令全中）
        raise EngineError(
            "bad-payload",
            "参数要写成一个 JSON 对象（{...}），这次收到的是 " + type(payload).__name__,
            "把参数包在花括号里：{\"text\": \"...\", \"outputPath\": \"...\"}")
    return command, payload


def main(argv):
    debug = "--debug" in argv
    try:
        command, payload = _parse_args(argv)
        handler = commands.TABLE.get(command)
        if handler is None:
            raise EngineError("unknown-command", "不认识命令：" + command, "可用命令：" + ", ".join(sorted(commands.TABLE)))
        # 学校画像兜底：老师没说的地方，按他平时习惯填（只补缺项，绝不覆盖本次明确要求）
        try:
            defaults.apply(payload, command)
        except Exception:
            pass  # 画像坏了不能拦住正事
        filled = payload.pop("_profileFilled", None)
        data = handler(payload)
        if filled and isinstance(data, dict):
            data["profileFilled"] = filled
        print(json.dumps({"ok": True, "command": command, "engine": __version__, "data": data}, ensure_ascii=False))
        return 0
    except EngineError as error:
        print(json.dumps({"ok": False, "error": error.to_dict()}, ensure_ascii=False))
        return 2
    except Exception as error:  # 非预期错误也要是可解析的 JSON
        if debug:
            traceback.print_exc(file=sys.stderr)
        print(json.dumps({
            "ok": False,
            "error": {"code": "internal", "message": "引擎内部错误：" + str(error)},
        }, ensure_ascii=False))
        return 1


if __name__ == "__main__":
    sys.exit(main(sys.argv))

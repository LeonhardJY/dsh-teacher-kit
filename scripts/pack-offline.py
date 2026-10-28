#!/usr/bin/env python3
"""把「离线一键安装包」打成一个 zip（给没装 Node、也不方便连 GitHub 的电脑用）。

    python scripts/pack-offline.py            # 输出 dist/dsh-teacher-kit-<版本>-offline-win.zip
    python scripts/pack-offline.py --keep     # 顺便留下临时装配目录，方便检查

包结构：

    dsh-teacher-kit-<版本>/
      安装说明.txt          CRLF，Windows 记事本直接能看
      一键安装.cmd          双击 → 调 install.ps1 %*
      一键卸载.cmd          双击 → 调 install.ps1 -Uninstall %*
      install.ps1          纯 PowerShell，不需要 Node
      plugin/              插件本体（lib / client / engine + 三份说明）
"""
import json
import os
import shutil
import sys
import tempfile
import zipfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
INSTALLER = os.path.join(ROOT, "offline-installer")
INSTALLER_FILES = ["安装说明.txt", "一键安装.cmd", "一键卸载.cmd", "install.ps1"]
PAYLOAD_DIRS = ["lib", "client", "engine"]
PAYLOAD_FILES = ["package.json", "cordis.patch.yml", "README.md", "RELEASE.md", "LICENSE"]
SKIP_DIRS = {"__pycache__", ".pytest_cache"}
SKIP_SUFFIX = (".pyc", ".log")


def main() -> int:
    keep = "--keep" in sys.argv
    with open(os.path.join(ROOT, "package.json"), encoding="utf-8") as handle:
        version = json.load(handle)["version"]
    name = "dsh-teacher-kit-" + version
    staging = tempfile.mkdtemp(prefix="teacher-kit-offline-")
    top = os.path.join(staging, name)
    os.makedirs(top)

    for filename in INSTALLER_FILES:
        source = os.path.join(INSTALLER, filename)
        if not os.path.exists(source):
            print("缺少安装器文件：" + source, file=sys.stderr)
            return 1
        target = os.path.join(top, filename)
        with open(source, "rb") as handle:
            text = handle.read().decode("utf-8").replace("\r\n", "\n").replace("\n", "\r\n")
        with open(target, "wb") as handle:
            handle.write(text.encode("utf-8"))

    plugin = os.path.join(top, "plugin")
    os.makedirs(plugin)
    for dirname in PAYLOAD_DIRS:
        shutil.copytree(
            os.path.join(ROOT, dirname),
            os.path.join(plugin, dirname),
            ignore=shutil.ignore_patterns(*SKIP_DIRS),
        )
    for filename in PAYLOAD_FILES:
        source = os.path.join(ROOT, filename)
        if os.path.exists(source):
            shutil.copy2(source, os.path.join(plugin, filename))

    outdir = os.path.join(ROOT, "dist")
    os.makedirs(outdir, exist_ok=True)
    out = os.path.join(outdir, name + "-offline-win.zip")
    if os.path.exists(out):
        os.remove(out)
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED, compresslevel=9) as archive:
        for dirpath, dirnames, filenames in os.walk(top):
            dirnames[:] = sorted(d for d in dirnames if d not in SKIP_DIRS)
            for filename in sorted(filenames):
                if filename.endswith(SKIP_SUFFIX):
                    continue
                full = os.path.join(dirpath, filename)
                rel = os.path.relpath(full, staging).replace(os.sep, "/")
                archive.write(full, rel)

    with zipfile.ZipFile(out) as archive:
        entries = archive.namelist()
        broken = archive.testzip()
    print("输出：" + out)
    print("大小：" + str(os.path.getsize(out)) + " 字节")
    print("条目：" + str(len(entries)))
    print("完整性：" + ("OK" if broken is None else "损坏 " + broken))
    if keep:
        print("装配目录：" + top)
    else:
        shutil.rmtree(staging, ignore_errors=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

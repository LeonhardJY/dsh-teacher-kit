@echo off
chcp 65001 >nul
setlocal
set "PS=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
if not exist "%PS%" set "PS=powershell"
title 教师助手 安装程序
"%PS%" -NoProfile -ExecutionPolicy Bypass -File "%~dp0install.ps1" %*
if errorlevel 1 (
  echo.
  echo 安装脚本报错了，请把上面的文字截图发给技术支持。
  pause
)
endlocal

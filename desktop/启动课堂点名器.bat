@echo off
title 课堂随机点名器
cd /d "%~dp0server"
python server.py
if errorlevel 1 (
  echo.
  echo [启动失败] 请截图本窗口内容联系维护人。
  pause
)

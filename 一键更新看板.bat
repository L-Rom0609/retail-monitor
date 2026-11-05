@echo off
setlocal
title 即时零售看板 - 双端同步更新
cd /d "%~dp0"

set "PYCMD="

REM 1) py 启动器
py -3 -c "import sys" >nul 2>nul
if not errorlevel 1 set "PYCMD=py -3"

REM 2) PATH 里的 python
if not defined PYCMD (
  python -c "import sys" >nul 2>nul
  if not errorlevel 1 set "PYCMD=python"
)

REM 3) WorkBuddy 自带 Python
if not defined PYCMD (
  for /d %%D in ("%USERPROFILE%\.workbuddy\binaries\python\versions\*") do (
    if not defined PYCMD (
      if exist "%%~D\python.exe" set "PYCMD=%%~D\python.exe"
    )
  )
)

if not defined PYCMD (
  echo.
  echo [错误] 没有找到可用的 Python 运行环境
  echo 请安装 Python 3.8 以上版本后重试。
  echo.
  pause
  exit /b 1
)

echo ============================================
echo   即时零售看板 - 双端同步更新
echo   本机: %COMPUTERNAME%
echo   时间: %date% %time%
echo ============================================
echo.

%PYCMD% "脚本\sync.py" %*
set "EC=%errorlevel%"

echo.
if "%EC%"=="0" (
  echo [完成] 更新成功，线上看板已同步: https://l-rom0609.github.io/retail-monitor/
) else (
  echo [失败] 更新未完成，退出码 %EC%，请把上面的报错信息截图反馈。
)
echo.
pause
exit /b %EC%

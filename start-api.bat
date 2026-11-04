@echo off
chcp 936 >nul 2>nul
setlocal enabledelayedexpansion

REM ============================================================
REM  文档工作台智能体 - 一键启动（FastAPI 接口 / 给小程序、对外用）
REM  用法：双击本文件即可；关闭窗口 = 停止服务
REM ------------------------------------------------------------
REM  【后续优化就改这里】
REM   1) 项目换位置  -> 改 PROJECT_DIR
REM   2) 换 Python   -> 改 PYTHON_EXE
REM   3) 换端口      -> 改 PORT
REM   4) 允许外部访问 -> HOST 改成 0.0.0.0（局域网/穿透需要）
REM ============================================================

set "PROJECT_DIR=E:\Vibecoding\doc-agent-workbench"
set "PYTHON_EXE=D:\Anaconda\envs\llm_course\python.exe"
set "PORT=8000"
set "HOST=127.0.0.1"
set "AUTO_OPEN_BROWSER=1"
set "BROWSER_DELAY=5"

title 文档工作台智能体 - FastAPI :%PORT%

echo ============================================================
echo    文档工作台智能体   接口服务（FastAPI）
echo ------------------------------------------------------------
echo    项目目录 : %PROJECT_DIR%
echo    监听地址 : %HOST%:%PORT%
echo    接口文档 : http://%HOST%:%PORT%/docs
echo ============================================================
echo.

REM ---------- 1. 环境检查 ----------
if not exist "%PROJECT_DIR%\server.py" (
    echo [错误] 找不到 server.py：%PROJECT_DIR%
    echo        请检查 PROJECT_DIR 配置是否正确。
    goto :error
)
if not exist "%PYTHON_EXE%" (
    echo [错误] 找不到 Python：%PYTHON_EXE%
    echo        请确认 Anaconda 环境 llm_course 已安装。
    goto :error
)

cd /d "%PROJECT_DIR%"

REM ---------- 2. 端口占用自动清理 ----------
set "OLD_PID="
for /f "tokens=5" %%P in ('netstat -ano ^| findstr /C:":%PORT% " ^| findstr "LISTENING"') do set "OLD_PID=%%P"
if defined OLD_PID (
    echo [提示] 端口 %PORT% 已被进程 !OLD_PID! 占用，正在结束它...
    taskkill /F /PID !OLD_PID! >nul 2>&1
    timeout /t 1 >nul
    echo [完成] 旧进程已清理。
    echo.
)

REM ---------- 3. 延迟自动打开浏览器 ----------
if "%AUTO_OPEN_BROWSER%"=="1" (
    start "" /min cmd /c "timeout /t %BROWSER_DELAY% >nul & start http://%HOST%:%PORT%"
)

REM ---------- 4. 启动服务 ----------
echo [启动] 正在启动 FastAPI，请稍候...
echo [访问] http://%HOST%:%PORT%
echo [停止] 按 Ctrl+C，或直接关闭本窗口
echo ------------------------------------------------------------
echo.

"%PYTHON_EXE%" -m uvicorn server:app --host %HOST% --port %PORT%

echo.
echo ------------------------------------------------------------
echo [结束] 服务已退出。
pause
exit /b 0

:error
echo.
pause
exit /b 1

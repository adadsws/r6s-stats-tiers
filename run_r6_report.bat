@echo off
setlocal
set "PYTHONUTF8=1"
set "PYTHONIOENCODING=utf-8"
chcp 65001 >nul
cd /d "%~dp0"

where python >nul 2>nul
if errorlevel 1 (
    where py >nul 2>nul
    if errorlevel 1 (
        echo [错误] 未找到 Python 3，请先安装 Python 3.9 或更高版本。
        pause
        exit /b 1
    )
    set "PYTHON_CMD=py -3"
) else (
    set "PYTHON_CMD=python"
)

if not exist "%~dp0data\athieno\latest.json" (
    echo [错误] 缺少 data\athieno\latest.json。
    echo 请先使用项目 Skill 人工核对 Athieno 最新 Tier 视频并保存评分。
    pause
    exit /b 1
)

set "PYTHONPATH=%~dp0src"

echo [1/3] 获取灰机 Wiki 最新数据、图标和补丁...
%PYTHON_CMD% -m r6_report.collector --data-dir "%~dp0data" --archive-dir "%~dp0~archived\data-snapshots" --temp-dir "%~dp0~temp"
if errorlevel 1 goto :failed

echo [2/3] 生成基础统计工作簿...
%PYTHON_CMD% -m r6_report.operator_stats --data-dir "%~dp0data" --output "%~dp0data\r6_operator_stats.xlsx"
if errorlevel 1 goto :failed

echo [3/3] 生成五个中文榜单...
%PYTHON_CMD% -m r6_report.leaderboards --data-dir "%~dp0data" --input "%~dp0data\r6_operator_stats.xlsx" --output-dir "%~dp0output"
if errorlevel 1 goto :failed

echo.
echo 完成。榜单位于：%~dp0output
pause
exit /b 0

:failed
set "EXIT_CODE=%ERRORLEVEL%"
echo.
echo [错误] 生成失败，退出代码 %EXIT_CODE%。
pause
exit /b %EXIT_CODE%

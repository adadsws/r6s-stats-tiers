@echo off
setlocal
set "PYTHONUTF8=1"
set "PYTHONIOENCODING=utf-8"
chcp 65001 >nul
for %%I in ("%~dp0..") do set "PROJECT_ROOT=%%~fI"
cd /d "%PROJECT_ROOT%"

echo [说明] 此脚本仅用于重新采集数据并重建榜单。
echo [说明] 如果只需查看或下载榜单，请直接使用 docs\outputs 中的成品。
echo.

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

if not exist "%PROJECT_ROOT%\data\athieno\latest.json" (
    echo [错误] 缺少 data\athieno\latest.json。
    echo 请先使用项目 Skill 人工核对 Athieno 最新 Tier 视频并保存评分。
    pause
    exit /b 1
)

set "PYTHONPATH=%PROJECT_ROOT%\src"

echo [1/3] 获取灰机 Wiki 最新数据、图标和补丁...
%PYTHON_CMD% -m r6_report.collector --data-dir "%PROJECT_ROOT%\data" --archive-dir "%PROJECT_ROOT%\~archive\data-snapshots" --temp-dir "%PROJECT_ROOT%\~temp"
if errorlevel 1 goto :failed

echo [2/3] 生成基础统计工作簿...
%PYTHON_CMD% -m r6_report.operator_stats --data-dir "%PROJECT_ROOT%\data" --output "%PROJECT_ROOT%\~temp\r6_operator_stats.xlsx"
if errorlevel 1 goto :failed

echo [3/3] 生成五个中文榜单...
%PYTHON_CMD% -m r6_report.leaderboards --data-dir "%PROJECT_ROOT%\data" --input "%PROJECT_ROOT%\~temp\r6_operator_stats.xlsx" --output-dir "%PROJECT_ROOT%\~outputs"
if errorlevel 1 goto :failed

echo.
echo 完成。榜单位于：%PROJECT_ROOT%\~outputs
pause
exit /b 0

:failed
set "EXIT_CODE=%ERRORLEVEL%"
echo.
echo [错误] 生成失败，退出代码 %EXIT_CODE%。
pause
exit /b %EXIT_CODE%

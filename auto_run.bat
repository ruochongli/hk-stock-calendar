@echo off
setlocal enabledelayedexpansion
chcp 65001 >nul
cd /d "%~dp0"

set LOGFILE=%~dp0logs\auto_run_%date:~0,4%%date:~5,2%%date:~8,2%.log

:: 记录开始时间
echo [%date% %time%] ========================================== >> "%LOGFILE%"
echo [%date% %time%]  自动运行港股公告爬虫 >> "%LOGFILE%"
echo [%date% %time%] ========================================== >> "%LOGFILE%"

:: 运行爬虫
python src/crawler.py >> "%LOGFILE%" 2>&1
if %errorlevel% neq 0 (
    echo [%date% %time%] [错误] 爬虫运行失败 >> "%LOGFILE%"
    exit /b 1
)

echo [%date% %time%] 爬虫运行完成，开始部署... >> "%LOGFILE%"

:: 检查 133 招商局基金公告并发送邮件通知
python scripts/notify_stock_133.py >> "%LOGFILE%" 2>&1

:: 推送到 GitHub Pages
git add index.html notices/ data/announcements.json
git commit -m "update: 自动更新港股公告日历 %date% %time%" >> "%LOGFILE%" 2>&1
git push origin main >> "%LOGFILE%" 2>&1

if %errorlevel% equ 0 (
    echo [%date% %time%] [成功] 已推送到 GitHub Pages >> "%LOGFILE%"
) else (
    echo [%date% %time%] [警告] 推送失败，可能是无更新或网络问题 >> "%LOGFILE%"
)

echo [%date% %time%] ========================================== >> "%LOGFILE%"
echo. >> "%LOGFILE%"

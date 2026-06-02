@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo ==========================================
echo  正在运行港股公告爬虫...
echo ==========================================
python src/crawler.py

echo.
echo ==========================================
echo  爬虫运行完成。
echo.
echo  本地预览: 双击 index.html
echo  网页部署: 双击 deploy.bat
echo ==========================================
pause

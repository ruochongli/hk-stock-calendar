@echo off
chcp 65001 >nul
echo ==========================================
echo   港股公告日历 - 创建每日自动更新任务
echo ==========================================
echo.

schtasks /create /tn "HK-Stock-Calendar-Auto-Update" /tr "C:\Users\ellenli\stock_crawler\auto_run.bat" /sc daily /st 09:00 /np /rl LIMITED

if %errorlevel% equ 0 (
    echo.
    echo [成功] 定时任务已创建！
    echo 每天上午 9:00 自动运行爬虫并推送到 GitHub
    echo.
    echo 可在 任务计划程序 中查看或修改：
    echo   开始菜单 → 任务计划程序
echo   任务计划程序库 → HK-Stock-Calendar-Auto-Update
) else (
    echo.
    echo [失败] 请确保以管理员身份运行此脚本
)

echo.
pause

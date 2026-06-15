@echo off
chcp 65001 >nul
echo ==========================================
echo   港股公告日历 - Pages 健康检查定时任务
echo ==========================================
echo.
echo 任务：每天 3:00 检查 GitHub Pages 是否可访问
echo       异常时自动重推并弹窗通知
echo.

schtasks /create /tn "HK-Stock-Calendar-Pages-Health" /tr "C:\Users\ellenli\stock_crawler\check_pages.bat" /sc daily /st 03:00 /np /rl LIMITED

if %errorlevel% equ 0 (
    echo.
    echo [成功] 定时任务已创建！
    echo 每天 3:00 自动检查：https://ruochongli.github.io/hk-stock-calendar/
    echo.
    echo 可在 任务计划程序 中查看或修改：
    echo   开始菜单 → 任务计划程序
    echo   任务计划程序库 → HK-Stock-Calendar-Pages-Health
) else (
    echo.
    echo [失败] 请确保以管理员身份运行此脚本
)

echo.
pause

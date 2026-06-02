@echo off
setlocal enabledelayedexpansion
chcp 65001 >nul
cd /d "%~dp0"

echo ==========================================
echo   港股公告日历 - GitHub Pages 自动部署
echo ==========================================
echo.

:: 检查 git
where git >nul 2>nul
if %errorlevel% neq 0 (
    echo [错误] 未找到 Git，请先安装 Git。
    echo 下载地址: https://git-scm.com/download/win
    pause
    exit /b 1
)

:: 检查是否有 remote
git remote get-url origin >nul 2>nul
if %errorlevel% neq 0 (
    echo [首次部署] 尚未关联 GitHub 仓库。
    echo.
    echo 请按以下步骤操作：
    echo 1. 访问 https://github.com/new 创建新仓库
    echo    建议仓库名: hk-stock-calendar
    echo 2. 不要勾选 "Initialize this repository with a README"
    echo 3. 创建后，复制仓库的 HTTPS 链接
    echo    格式类似: https://github.com/你的用户名/hk-stock-calendar.git
    echo.
    set /p repo_url="请粘贴仓库链接: "
    
    if "!repo_url!"=="" (
        echo [取消] 未输入链接，部署已取消。
        pause
        exit /b 0
    )
    
    git remote add origin !repo_url!
    echo [成功] 已关联远程仓库。
    echo.
)

:: 获取远程仓库信息，生成 Pages 链接
for /f "tokens=*" %%a in ('git remote get-url origin 2^>nul') do set remote_url=%%a

:: 提取用户名和仓库名
set pages_url=无法自动推断，请手动查看 GitHub 仓库 Settings -> Pages
for /f "tokens=3,4 delims=/" %%a in ("%remote_url%") do (
    set gh_user=%%a
    set gh_repo=%%b
    set gh_repo=!gh_repo:.git=! 
    set pages_url=https://%%a.github.io/!gh_repo!
)

:: 确保只推送 index.html
git checkout -b main 2>nul
git add index.html
git commit -m "update: 港股公告日历 %date% %time%" >nul 2>nul

:: 推送
git push -u origin main --force
echo.

if %errorlevel% equ 0 (
    echo ==========================================
    echo   [部署成功] 
    echo.
    echo   网页链接: %pages_url%
    echo.
    echo   首次部署后，请访问 GitHub 仓库：
    echo   Settings -> Pages -> Source
echo   确认已选择 "Deploy from a branch" + "main"
    echo   等待约 1 分钟后，链接即可访问。
    echo ==========================================
) else (
    echo [部署失败] 请检查网络或 GitHub 凭据。
    echo 提示: 如果提示需要登录，请使用 Personal Access Token
    echo       或在 Git 中配置好 SSH 密钥。
)

echo.
pause

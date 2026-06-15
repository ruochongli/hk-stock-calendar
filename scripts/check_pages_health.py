#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
GitHub Pages 健康检查 Loop

- 检查 Pages 是否可访问（HTTP 200）
- 如果异常，尝试重推现有 index.html / notices/ / data/announcements.json
- 记录日志到 logs/pages_health_YYYYMMDD.log
- 异常或修复失败时发送 Windows 通知
"""

import datetime
import subprocess
import sys
from pathlib import Path

import requests

PROJECT_ROOT = Path(__file__).resolve().parent.parent
LOGS_DIR = PROJECT_ROOT / "logs"
PAGES_URL = "https://ruochongli.github.io/hk-stock-calendar/"
TIMEOUT = 30


def log(msg: str) -> None:
    now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{now}] {msg}"
    print(line)
    LOGS_DIR.mkdir(exist_ok=True)
    today = datetime.date.today().strftime("%Y%m%d")
    with open(LOGS_DIR / f"pages_health_{today}.log", "a", encoding="utf-8") as f:
        f.write(line + "\n")


def _send_toast(title: str, message: str) -> None:
    """Windows 气泡通知，作为邮件失败时的 fallback。"""
    try:
        safe_title = title.replace("'", "''")
        safe_message = message.replace("'", "''")
        cmd = (
            "Add-Type -AssemblyName System.Windows.Forms; "
            "$n = [System.Windows.Forms.NotifyIcon]::new(); "
            "$n.Icon = [System.Drawing.SystemIcons]::Information; "
            "$n.Visible = $true; "
            f"$n.ShowBalloonTip(5000, '{safe_title}', '{safe_message}', "
            "[System.Windows.Forms.ToolTipIcon]::Info)"
        )
        subprocess.run(
            ["powershell", "-Command", cmd],
            cwd=PROJECT_ROOT,
            capture_output=True,
            text=True,
            check=False,
            timeout=10,
        )
    except Exception:
        pass


def notify(title: str, message: str) -> None:
    """通过 Outlook 发送邮件通知；如果 Outlook 不可用，降级为 Windows 通知。"""
    email_sent = False
    try:
        safe_title = title.replace("'", "''")
        safe_message = message.replace("'", "''").replace("\n", "`n")
        cmd = (
            "$ErrorActionPreference = 'Stop'; "
            "$outlook = New-Object -ComObject Outlook.Application; "
            "$mail = $outlook.CreateItem(0); "
            "$mail.To = 'ellenli@tfisec.com'; "
            f"$mail.Subject = '[港股公告日历] {safe_title}'; "
            f"$mail.Body = '{safe_message}'; "
            "$mail.Send()"
        )
        result = subprocess.run(
            ["powershell", "-Command", cmd],
            cwd=PROJECT_ROOT,
            capture_output=True,
            text=True,
            check=False,
            timeout=30,
        )
        if result.returncode == 0:
            log("邮件通知已发送")
            email_sent = True
        else:
            log(f"邮件发送失败: {result.stderr.strip()}")
    except Exception as e:
        log(f"邮件发送失败: {e}")

    if not email_sent:
        _send_toast(title, message)
        log("已降级为桌面通知")


def check_pages() -> tuple[int | None, str]:
    try:
        r = requests.get(PAGES_URL, timeout=TIMEOUT)
        return r.status_code, r.reason
    except requests.RequestException as e:
        return None, str(e)


def redeploy() -> bool:
    log("开始重推 GitHub Pages...")
    try:
        result = subprocess.run(
            ["git", "add", "index.html", "notices/", "data/announcements.json"],
            cwd=PROJECT_ROOT,
            capture_output=True,
            text=True,
        )
        if result.returncode != 0:
            log(f"git add 失败: {result.stderr.strip()}")
            return False

        now = datetime.datetime.now().isoformat()
        result = subprocess.run(
            ["git", "commit", "-m", f"health check redeploy {now}"],
            cwd=PROJECT_ROOT,
            capture_output=True,
            text=True,
        )
        # 没有变更时 commit 会失败，但不影响后续 push
        if result.returncode != 0 and "nothing to commit" not in result.stderr.lower():
            log(f"git commit 失败: {result.stderr.strip()}")
            return False

        result = subprocess.run(
            ["git", "push", "origin", "main"],
            cwd=PROJECT_ROOT,
            capture_output=True,
            text=True,
        )
        if result.returncode != 0:
            log(f"git push 失败: {result.stderr.strip()}")
            return False

        log("重推成功")
        return True
    except Exception as e:
        log(f"重推异常: {e}")
        return False


def main() -> int:
    log(f"开始检查 Pages: {PAGES_URL}")
    status, reason = check_pages()

    if status == 200:
        log(f"Pages 健康: HTTP {status}")
        return 0

    log(f"Pages 异常: HTTP {status}, reason={reason}")
    success = redeploy()

    if success:
        log("已尝试修复，GitHub Pages 刷新约需 1-3 分钟")
        notify("港股公告日历", "Pages 曾异常，已重推，请稍后检查")
    else:
        log("自动修复失败，需要人工介入")
        notify("港股公告日历", "Pages 异常且自动修复失败，请人工检查")

    return 1


if __name__ == "__main__":
    sys.exit(main())

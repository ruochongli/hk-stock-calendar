#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
133 招商局基金公告邮件通知

- 读取 data/announcements.json
- 筛选 stock_code 为 00133 的公告
- 通过 Outlook 发送邮件到 ellenli@tfisec.com

用法：
    python scripts/notify_stock_133.py [--date YYYY-MM-DD]

不带 --date 时，默认检查今天（运行日期）的公告。
"""

import argparse
import datetime
import json
import subprocess
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent
DATA_PATH = PROJECT_ROOT / "data" / "announcements.json"
RECIPIENT = "ellenli@tfisec.com"
TARGET_CODE = "00133"


def log(msg: str) -> None:
    now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"[{now}] {msg}")


def send_email(subject: str, body: str) -> bool:
    """通过 Outlook COM 发送邮件。"""
    try:
        safe_subject = subject.replace("'", "''")
        safe_body = body.replace("'", "''").replace("\n", "`n")
        cmd = (
            "$ErrorActionPreference = 'Stop'; "
            "$outlook = New-Object -ComObject Outlook.Application; "
            "$mail = $outlook.CreateItem(0); "
            f"$mail.To = '{RECIPIENT}'; "
            f"$mail.Subject = '{safe_subject}'; "
            f"$mail.Body = '{safe_body}'; "
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
            return True
        else:
            log(f"邮件发送失败: {result.stderr.strip()}")
            return False
    except Exception as e:
        log(f"邮件发送失败: {e}")
        return False


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--date", help="要检查的日期，格式 YYYY-MM-DD，默认今天")
    args = parser.parse_args()

    target_date = args.date or datetime.date.today().isoformat()
    log(f"检查 {target_date} 的 {TARGET_CODE} 公告...")

    if not DATA_PATH.exists():
        log(f"未找到公告数据: {DATA_PATH}")
        return 1

    with open(DATA_PATH, "r", encoding="utf-8") as f:
        announcements = json.load(f)

    matched = [
        a for a in announcements
        if a.get("stock_code") == TARGET_CODE and a.get("notice_date") == target_date
    ]

    if not matched:
        log(f"{target_date} 没有 {TARGET_CODE} 公告")
        return 0

    log(f"发现 {len(matched)} 条 {TARGET_CODE} 公告，准备发送邮件...")

    lines = [
        f"【{a['stock_code']}】【{a['stock_name']}】【{a['notice_date']}】发布公告【{a['title']}】"
        for a in matched
    ]
    body = "\n".join(lines)
    subject = f"[港股公告] 招商局基金({TARGET_CODE}) {target_date} 共{len(matched)}条公告"

    if send_email(subject, body):
        return 0
    return 1


if __name__ == "__main__":
    sys.exit(main())

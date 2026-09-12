#!/usr/bin/env python3
"""下载最新目标邮件，供人工查看；月末汇总请使用 monthly_summary.py。"""

from __future__ import annotations

import argparse
import os
import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from pipeline.monthly_summary import fetch_latest_target_attachment_any, save_latest_mail  # noqa: E402


def main() -> int:
    parser = argparse.ArgumentParser(description="下载最新物料情况邮件附件和 HTML")
    parser.add_argument("--output-dir", default="", help="输出目录，默认读取 LATEST_MAIL_DIR")
    parser.add_argument("--limit", type=int, default=100, help="检查最近多少封邮件，默认 100")
    parser.add_argument("--subject", default="", help="目标邮件主题关键词")
    parser.add_argument("--attachment-pattern", default="", help="目标附件匹配模式")
    args = parser.parse_args()

    try:
        from dotenv import load_dotenv
        load_dotenv(REPO_ROOT / ".env")
    except ImportError:
        pass

    output_value = args.output_dir or os.getenv("LATEST_MAIL_DIR", "").strip()
    if not output_value:
        print("❌ 未配置 LATEST_MAIL_DIR，也未提供 --output-dir")
        return 2
    output_dir = Path(output_value).expanduser()
    subject = args.subject or os.getenv("TARGET_MAIL_SUBJECT", "物料情况和Excel文件")
    attachment_pattern = args.attachment_pattern or os.getenv("TARGET_ATTACHMENT_PATTERN", "总库存*.xlsx")
    user = os.getenv("EMAIL_ADDRESS_QQ", "").strip()
    password = os.getenv("EMAIL_PASSWORD_QQ") or os.getenv("EMAIL_PASSWOR_QQ", "")
    if not user or not password:
        print("❌ 缺少 EMAIL_ADDRESS_QQ / EMAIL_PASSWORD_QQ")
        return 2
    try:
        candidate = fetch_latest_target_attachment_any(
            server=os.getenv("IMAP_SERVER", "imap.qq.com"),
            user=user,
            password=password,
            mailbox=os.getenv("IMAP_MAILBOX", "INBOX"),
            subject_keyword=subject,
            attachment_pattern=attachment_pattern,
            limit=max(1, args.limit),
        )
        attachment, html = save_latest_mail(candidate=candidate, output_dir=output_dir)
        print(f"✅ 最新邮件: {candidate.received_at.isoformat()} | {candidate.subject}")
        print(f"✅ Excel: {attachment}")
        if html:
            print(f"✅ HTML: {html}")
        return 0
    except Exception as exc:
        print(f"❌ 最新邮件下载失败: {exc}")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())

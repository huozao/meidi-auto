#!/usr/bin/env python3
"""下载月末目标邮件并生成物料月度分析 Excel。"""

from __future__ import annotations

import argparse
import os
import sys
from datetime import datetime
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from pipeline.monthly_summary import (  # noqa: E402
    TZ_SHANGHAI,
    aggregate_sources,
    build_report_frames,
    append_manifest_record,
    fetch_latest_target_attachment,
    month_cutoff,
    month_token,
    save_month_snapshot,
    MonthlySource,
    read_month_workbook,
    scan_month_files,
    write_report,
)


def _load_local_env() -> None:
    try:
        from dotenv import load_dotenv
    except ImportError:
        return
    load_dotenv(REPO_ROOT / ".env")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="按月末邮件和月末文件生成物料汇总 Excel")
    parser.add_argument("--month", help="下载指定月份的月末邮件，格式 YYYY-MM")
    parser.add_argument("--aggregate-only", action="store_true", help="不访问邮箱，仅汇总已有文件")
    parser.add_argument("--dry-run", action="store_true", help="只检查邮件/文件和汇总范围，不写入文件")
    parser.add_argument("--archive-dir", default="", help="月度汇总根目录，默认读取 MONTHLY_ARCHIVE_DIR")
    parser.add_argument("--auto-archive-dir", default="", help="自动月末归档目录，默认根目录/自动月末归档")
    parser.add_argument("--raw-snapshot-dir", default="", help="原始邮件快照目录，默认根目录/原始月末快照")
    parser.add_argument("--report-dir", default="", help="报告目录，默认读取 MONTHLY_REPORT_DIR 或根目录/分析结果")
    parser.add_argument("--mailbox", default="", help="IMAP 邮箱目录，默认读取 IMAP_MAILBOX 或 INBOX")
    parser.add_argument("--subject", default="", help="目标邮件主题关键词")
    parser.add_argument("--attachment-pattern", default="", help="目标附件匹配模式")
    return parser.parse_args()


def _required_env(name: str) -> str:
    value = os.getenv(name, "").strip()
    if not value:
        raise RuntimeError(f"缺少环境变量: {name}")
    return value


def _report_path(report_dir: Path, sources) -> Path:
    first = sources[0].month.replace("-", "")[2:]
    last = sources[-1].month.replace("-", "")[2:]
    timestamp = datetime.now(TZ_SHANGHAI).strftime("%Y%m%d_%H%M%S")
    return report_dir / f"月度物料分析_{first}-{last}_{timestamp}.xlsx"


def main() -> int:
    _load_local_env()
    args = parse_args()
    if not args.aggregate_only and not args.month:
        print("❌ 下载模式必须提供 --month YYYY-MM；仅汇总已有文件请使用 --aggregate-only")
        return 2

    root_value = args.archive_dir or os.getenv("MONTHLY_ARCHIVE_DIR", "").strip()
    if not root_value:
        print("❌ 未配置 MONTHLY_ARCHIVE_DIR，也未提供 --archive-dir")
        return 2
    root = Path(root_value).expanduser()
    auto_dir = Path(args.auto_archive_dir or os.getenv("MONTHLY_AUTO_ARCHIVE_DIR", root / "自动月末归档")).expanduser()
    raw_dir = Path(args.raw_snapshot_dir or os.getenv("MONTHLY_RAW_SNAPSHOT_DIR", root / "原始月末快照")).expanduser()
    report_dir = Path(args.report_dir or os.getenv("MONTHLY_REPORT_DIR", root / "分析结果")).expanduser()

    try:
        if args.month and not args.aggregate_only:
            server = os.getenv("IMAP_SERVER", "imap.qq.com")
            mailbox = args.mailbox or os.getenv("IMAP_MAILBOX", "INBOX")
            subject = args.subject or os.getenv("TARGET_MAIL_SUBJECT", "物料情况和Excel文件")
            attachment_pattern = args.attachment_pattern or os.getenv("TARGET_ATTACHMENT_PATTERN", "总库存*.xlsx")
            candidate = fetch_latest_target_attachment(
                server=server,
                user=_required_env("EMAIL_ADDRESS_QQ"),
                password=os.getenv("EMAIL_PASSWORD_QQ") or _required_env("EMAIL_PASSWOR_QQ"),
                mailbox=mailbox,
                month=args.month,
                subject_keyword=subject,
                attachment_pattern=attachment_pattern,
            )
            cutoff = month_cutoff(args.month)
            print(f"✅ 目标邮件: {candidate.received_at.isoformat()} | {candidate.subject}")
            print(f"✅ 附件: {candidate.filename} | {len(candidate.payload)} bytes | 截止: {cutoff.isoformat()}")
            if not args.dry_run:
                saved = save_month_snapshot(
                    candidate=candidate,
                    month=args.month,
                    auto_archive_dir=auto_dir,
                    raw_snapshot_dir=raw_dir,
                )
                quality = read_month_workbook(MonthlySource(args.month, saved, "automatic")).quality
                raw_candidates = sorted(raw_dir.glob(f"{month_token(args.month)}_*.xlsx"), key=lambda path: path.stat().st_mtime)
                raw_path = raw_candidates[-1] if raw_candidates else raw_dir / saved.name
                append_manifest_record(
                    manifest_path=raw_dir / "manifest.jsonl",
                    candidate=candidate,
                    month=args.month,
                    raw_path=raw_path,
                    canonical_path=saved,
                    quality=quality,
                )
                print(f"✅ 自动月末文件: {saved}")
        sources = scan_month_files(root, auto_dir)
        if not sources:
            raise ValueError(f"未找到月末文件: {root}")
        print(f"✅ 扫描到 {len(sources)} 个月份: {sources[0].month} ~ {sources[-1].month}")
        result = aggregate_sources(sources)
        errors = [item for item in result.quality if item.get("status") == "error"]
        warnings = [item for item in result.quality if item.get("status") == "warning"]
        print(f"✅ 物料月度记录: {len(result.records)} | 错误: {len(errors)} | 警告: {len(warnings)}")
        for item in errors:
            print(f"❌ {item.get('month')}: {item.get('error', '工作簿读取失败')}")
        for item in warnings:
            print(f"⚠️ {item.get('month')}: {item.get('detail_warning', item.get('error', '质量警告'))}")
        if args.dry_run:
            print("ℹ️ dry-run：未写入自动归档和报告")
            return 0
        output = _report_path(report_dir, sources)
        write_report(build_report_frames(result), output)
        print(f"✅ 汇总报告: {output}")
        return 0
    except Exception as exc:
        print(f"❌ 月度汇总失败: {exc}")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())

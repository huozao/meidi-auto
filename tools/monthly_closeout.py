#!/usr/bin/env python3
"""月初月末核验：核对上一月最后一天每日归档与月末邮件附件后生成报告。"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import smtplib
import sys
import tempfile
from datetime import date, datetime, timedelta
from email.mime.text import MIMEText
from pathlib import Path

import openpyxl
from dotenv import load_dotenv

REPO_ROOT = Path(__file__).resolve().parent.parent
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from pipeline.archive import WebDavStore  # noqa: E402
from pipeline.monthly_summary import (  # noqa: E402
    TZ_SHANGHAI,
    MonthlySource,
    aggregate_sources,
    append_manifest_record,
    build_report_frames,
    fetch_latest_target_attachment,
    month_cutoff,
    month_token,
    read_month_workbook,
    save_month_snapshot,
    scan_month_files,
    write_report,
)


def previous_month(today: date) -> str:
    first = today.replace(day=1)
    previous = first - timedelta(days=1)
    return previous.strftime("%Y-%m")


def _value(value: object) -> object:
    if isinstance(value, datetime):
        return value.isoformat()
    if isinstance(value, date):
        return value.isoformat()
    if value is None:
        return None
    return str(value)


def workbook_fingerprint(path: Path) -> str:
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        payload = []
        for ws in wb.worksheets:
            rows = [[_value(cell) for cell in row] for row in ws.iter_rows()]
            payload.append([ws.title, rows])
        encoded = json.dumps(payload, ensure_ascii=False, separators=(",", ":")).encode("utf-8")
        return hashlib.sha256(encoded).hexdigest()
    finally:
        wb.close()


def _required(name: str) -> str:
    value = os.getenv(name, "").strip()
    if not value:
        raise RuntimeError(f"缺少环境变量: {name}")
    return value


def _send_alert(month: str, daily: Path, downloaded: str, daily_fp: str, mail_fp: str) -> None:
    user = _required("EMAIL_ADDRESS_QQ")
    password = os.getenv("EMAIL_PASSWORD_QQ") or _required("EMAIL_PASSWOR_QQ")
    recipients = [item.strip() for item in re.split(r"[,;]", os.getenv("MONTHLY_ALERT_RECIPIENTS", os.getenv("RECIPIENT_EMAILS", ""))) if item.strip()]
    if not recipients:
        raise RuntimeError("未设置 MONTHLY_ALERT_RECIPIENTS 或 RECIPIENT_EMAILS")
    body = (f"{month} 月末文件核验未通过，请人工核验。\n"
            f"每日归档: {daily}\n下载附件: {downloaded}\n"
            f"每日归档语义指纹: {daily_fp}\n下载附件语义指纹: {mail_fp}\n")
    msg = MIMEText(body, "plain", "utf-8")
    msg["From"] = user
    msg["To"] = ", ".join(recipients)
    msg["Subject"] = f"美的月末文件核验异常 - {month}"
    with smtplib.SMTP("smtp.qq.com", 587, timeout=30) as server:
        server.starttls()
        server.login(user, password)
        server.send_message(msg)
    print(f"⚠️ 核验不一致，已发送人工核验通知（{len(recipients)} 位收件人）")


def _download_remote_monthly(store: WebDavStore, remote_root: str, temp_root: Path) -> tuple[Path, Path, Path]:
    monthly_root = temp_root / "月度汇总"
    auto_dir = monthly_root / "自动月末归档"
    raw_dir = monthly_root / "原始月末快照"
    report_dir = monthly_root / "分析结果"
    auto_dir.mkdir(parents=True)
    raw_dir.mkdir()
    report_dir.mkdir()
    for remote_dir, local_dir in ((f"{remote_root}/自动月末归档", auto_dir), (remote_root, monthly_root)):
        for name in store.list_files(remote_dir):
            if name.endswith(".xlsx") and (name.endswith("月底.xlsx") or remote_dir.endswith("自动月末归档")):
                local_dir.joinpath(name).write_bytes(store.get(f"{remote_dir}/{name}"))
    return monthly_root, auto_dir, report_dir


def _find_daily_file(root: Path, month: str) -> Path:
    folder = month.replace("-", "")
    last_day = month_cutoff(month).strftime("%Y%m%d")
    directory = root / folder
    candidates = []
    if directory.exists():
        for path in directory.glob("*.xlsx"):
            if last_day in path.name:
                candidates.append(path)
    if not candidates:
        raise FileNotFoundError(f"未找到 {month} 月末当天（{last_day}）的每日归档 Excel: {directory}")
    return max(candidates, key=lambda path: path.stat().st_mtime)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="核验月末每日归档与邮件附件并生成月度报告")
    parser.add_argument("--month", help="目标月份 YYYY-MM；默认按上海时间取上月")
    parser.add_argument("--dry-run", action="store_true", help="只核验，不写报告、不归档、不发邮件")
    return parser.parse_args()


def main() -> int:
    try:
        load_dotenv(REPO_ROOT / ".env")
        args = parse_args()
        month = args.month or previous_month(datetime.now(TZ_SHANGHAI).date())
        cutoff = month_cutoff(month)
        local_daily_root = os.getenv("DAILY_ATTACHMENT_ARCHIVE_DIR", "").strip()
        dav_base = os.getenv("NUTSTORE_WEBDAV_URL", "").strip()
        dav_user = os.getenv("NUTSTORE_WEBDAV_USER", "").strip()
        dav_password = os.getenv("NUTSTORE_WEBDAV_APP_PASSWORD", "").strip()
        dav_daily_root = os.getenv("NUTSTORE_REMOTE_DAILY_ARCHIVE_DIR", "").strip()
        dav_monthly_root = os.getenv("NUTSTORE_REMOTE_MONTHLY_ARCHIVE_DIR", "").strip()
        if not local_daily_root and not (dav_base and dav_user and dav_password and dav_daily_root and dav_monthly_root):
            raise RuntimeError("请配置本地 DAILY_ATTACHMENT_ARCHIVE_DIR，或完整配置 Nutstore WebDAV 环境变量")

        with tempfile.TemporaryDirectory(prefix="meidi-closeout-") as temp:
            temp_root = Path(temp)
            store = WebDavStore(dav_base, dav_user, dav_password) if dav_base and dav_user and dav_password else None
            if store:
                monthly_root, auto_dir, report_dir = _download_remote_monthly(store, dav_monthly_root, temp_root)
                daily_root = temp_root / "RE"
                daily_folder = month.replace("-", "")
                daily_dir = daily_root / daily_folder
                daily_dir.mkdir(parents=True)
                for name in store.list_files(f"{dav_daily_root}/{daily_folder}"):
                    if name.endswith(".xlsx"):
                        (daily_dir / name).write_bytes(store.get(f"{dav_daily_root}/{daily_folder}/{name}"))
            else:
                monthly_root = Path(_required("MONTHLY_ARCHIVE_DIR")).expanduser()
                auto_dir = Path(os.getenv("MONTHLY_AUTO_ARCHIVE_DIR", monthly_root / "自动月末归档")).expanduser()
                report_dir = Path(os.getenv("MONTHLY_REPORT_DIR", monthly_root / "分析结果")).expanduser()
                daily_root = Path(local_daily_root).expanduser()

            daily_path = _find_daily_file(daily_root, month)
            candidate = fetch_latest_target_attachment(
                server=os.getenv("IMAP_SERVER", "imap.qq.com"), user=_required("EMAIL_ADDRESS_QQ"),
                password=os.getenv("EMAIL_PASSWORD_QQ") or _required("EMAIL_PASSWOR_QQ"),
                mailbox=os.getenv("IMAP_MAILBOX", "INBOX"), month=month,
                subject_keyword=os.getenv("TARGET_MAIL_SUBJECT", "物料情况和Excel文件"),
                attachment_pattern=os.getenv("TARGET_ATTACHMENT_PATTERN", "总库存*.xlsx"),
            )
            mail_path = temp_root / candidate.filename
            mail_path.write_bytes(candidate.payload)
            daily_fp = workbook_fingerprint(daily_path)
            mail_fp = workbook_fingerprint(mail_path)
            print(f"✅ 月份: {month} | 截止: {cutoff.isoformat()}")
            print(f"✅ 每日归档: {daily_path.name} | 指纹: {daily_fp}")
            print(f"✅ 下载附件: {candidate.filename} | 指纹: {mail_fp}")
            if daily_fp != mail_fp:
                if args.dry_run:
                    print("⚠️ dry-run：发现不一致，未发送通知")
                    return 2
                _send_alert(month, daily_path, candidate.filename, daily_fp, mail_fp)
                return 2
            if args.dry_run:
                print("✅ dry-run：文件一致，未写入报告")
                return 0

            saved = save_month_snapshot(candidate=candidate, month=month, auto_archive_dir=auto_dir, raw_snapshot_dir=monthly_root / "原始月末快照")
            source_list = scan_month_files(monthly_root, auto_dir)
            result = aggregate_sources(source_list)
            output = report_dir / f"月度物料分析_{month_token(source_list[0].month)}-{month_token(source_list[-1].month)}_{datetime.now(TZ_SHANGHAI).strftime('%Y%m%d_%H%M%S')}.xlsx"
            write_report(build_report_frames(result), output)
            raw_candidates = sorted((monthly_root / "原始月末快照").glob(f"{month_token(month)}_*.xlsx"))
            raw_path = raw_candidates[-1] if raw_candidates else saved
            append_manifest_record(manifest_path=monthly_root / "原始月末快照" / "manifest.jsonl", candidate=candidate, month=month, raw_path=raw_path, canonical_path=saved, quality=read_month_workbook(MonthlySource(month, saved, "automatic")).quality)
            if store:
                store.put(f"{dav_monthly_root}/自动月末归档/{saved.name}", saved.read_bytes())
                store.put(f"{dav_monthly_root}/原始月末快照/{raw_path.name}", raw_path.read_bytes())
                store.put(f"{dav_monthly_root}/原始月末快照/manifest.jsonl", (monthly_root / "原始月末快照" / "manifest.jsonl").read_bytes())
                store.put(f"{dav_monthly_root}/分析结果/{output.name}", output.read_bytes())
                print(f"✅ 报告已回写坚果云: {dav_monthly_root}/分析结果/{output.name}")
            else:
                print(f"✅ 汇总报告: {output}")
            return 0
    except Exception as exc:
        print(f"❌ 月初月末核验失败: {exc}")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())

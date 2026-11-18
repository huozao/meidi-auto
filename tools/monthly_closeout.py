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
from pipeline.auto_list import generate_auto_list  # noqa: E402
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


def _fetch_optional_mail_candidate(month: str, cutoff: datetime) -> MailAttachmentCandidate | None:
    """尝试从 IMAP 检索目标月份月末邮件，作为审计与元数据参考；失败或未找到时不阻断归档。"""
    user = os.getenv("EMAIL_ADDRESS_QQ", "").strip()
    password = os.getenv("EMAIL_PASSWORD_QQ", "").strip() or os.getenv("EMAIL_PASSWOR_QQ", "").strip()
    if not (user and password):
        return None
    server = os.getenv("IMAP_SERVER", "imap.qq.com")
    mailbox = os.getenv("IMAP_MAILBOX", "INBOX")
    # 关键词优先读取配置；默认支持上游骏都仓、系统日报及历史格式
    subject_candidates = [s.strip() for s in os.getenv("TARGET_MAIL_SUBJECT", "合肥市和裕达,骏都对帐表,美的库存及出入库日报,物料情况").split(",") if s.strip()]
    attachment_pattern = os.getenv("TARGET_ATTACHMENT_PATTERN", "*.xlsx")

    for subject_kw in subject_candidates:
        try:
            return fetch_latest_target_attachment(
                server=server,
                user=user,
                password=password,
                mailbox=mailbox,
                month=month,
                subject_keyword=subject_kw,
                attachment_pattern=attachment_pattern,
            )
        except Exception:
            continue
    return None


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
                list_path = temp_root / "auto-list.xlsx"
                remote_list_path = os.getenv("NUTSTORE_REMOTE_AUTO_LIST_FILE", f"{dav_monthly_root}/自动预警清单/list.xlsx").strip()
                try:
                    list_path.write_bytes(store.get(remote_list_path))
                except RuntimeError:
                    print("ℹ️ 坚果云尚无上一版自动清单，将按历史月末数据首次生成。")
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
                list_path = Path(os.getenv("AUTO_LIST_FILE", monthly_root / "自动预警清单" / "list.xlsx")).expanduser()

            daily_path = _find_daily_file(daily_root, month)
            daily_source = MonthlySource(month, daily_path, "daily_archive")
            daily_result = read_month_workbook(daily_source)
            if daily_result.quality.get("status") == "error":
                raise ValueError(f"每日归档月末文件校验失败: {daily_result.quality.get('error')}")

            daily_payload = daily_path.read_bytes()
            daily_fp = workbook_fingerprint(daily_path)
            print(f"✅ 月份: {month} | 截止: {cutoff.isoformat()}")
            print(f"✅ 每日归档月末文件: {daily_path.name} | 指纹: {daily_fp} | 物料数: {daily_result.quality.get('material_count')}")

            mail_candidate = _fetch_optional_mail_candidate(month, cutoff)
            if mail_candidate:
                mail_path = temp_root / mail_candidate.filename
                mail_path.write_bytes(mail_candidate.payload)
                mail_fp = workbook_fingerprint(mail_path)
                print(f"ℹ️ 检索到月末参考邮件: {mail_candidate.subject} | 附件: {mail_candidate.filename} | 指纹: {mail_fp}")
            else:
                print("ℹ️ 未检索到对应月末邮件或跳过邮件检索，以每日归档处理后文件为准。")

            if args.dry_run:
                print("✅ dry-run：月末文件校验通过，未写入报告")
                return 0

            # 以处理后的每日归档总库存文件为正本生成月末快照
            snapshot_candidate = MailAttachmentCandidate(
                uid=mail_candidate.uid if mail_candidate else f"daily_archive_{daily_path.stat().st_mtime_ns}",
                received_at=mail_candidate.received_at if mail_candidate else datetime.fromtimestamp(daily_path.stat().st_mtime, tz=TZ_SHANGHAI),
                subject=mail_candidate.subject if mail_candidate else f"美的库存及出入库日报｜{month}",
                filename=daily_path.name,
                payload=daily_payload,
            )
            saved = save_month_snapshot(candidate=snapshot_candidate, month=month, auto_archive_dir=auto_dir, raw_snapshot_dir=monthly_root / "原始月末快照")
            source_list = scan_month_files(monthly_root, auto_dir)
            result = aggregate_sources(source_list)
            output = report_dir / f"月度物料分析_{month_token(source_list[0].month)}-{month_token(source_list[-1].month)}_{datetime.now(TZ_SHANGHAI).strftime('%Y%m%d_%H%M%S')}.xlsx"
            write_report(build_report_frames(result), output)
            remarks_source = list_path if list_path.exists() else REPO_ROOT / "script" / "data" / "list.xlsx"
            auto_list = generate_auto_list(result, month, list_path, remarks_source=remarks_source)
            # 当前清单固定覆盖，历史清单按月份留存；本地和 WebDAV 两种运行模式都保留。
            local_history_path = monthly_root / "自动清单历史" / f"{month_token(month)}.xlsx"
            local_history_path.parent.mkdir(parents=True, exist_ok=True)
            local_history_path.write_bytes(auto_list.path.read_bytes())
            raw_candidates = sorted((monthly_root / "原始月末快照").glob(f"{month_token(month)}_*.xlsx"))
            raw_path = raw_candidates[-1] if raw_candidates else saved
            append_manifest_record(manifest_path=monthly_root / "原始月末快照" / "manifest.jsonl", candidate=snapshot_candidate, month=month, raw_path=raw_path, canonical_path=saved, quality=read_month_workbook(MonthlySource(month, saved, "automatic")).quality)
            if store:
                store.put(f"{dav_monthly_root}/自动月末归档/{saved.name}", saved.read_bytes())
                store.put(f"{dav_monthly_root}/原始月末快照/{raw_path.name}", raw_path.read_bytes())
                store.put(f"{dav_monthly_root}/原始月末快照/manifest.jsonl", (monthly_root / "原始月末快照" / "manifest.jsonl").read_bytes())
                store.put(remote_list_path, auto_list.path.read_bytes())
                store.put(f"{dav_monthly_root}/自动清单历史/{month_token(month)}.xlsx", auto_list.path.read_bytes())
                store.put(f"{dav_monthly_root}/分析结果/{output.name}", output.read_bytes())
                print(f"✅ 自动清单: {auto_list.path} ({auto_list.material_count} 种物料)")
                print(f"✅ 报告已回写坚果云: {dav_monthly_root}/分析结果/{output.name}")
            else:
                print(f"✅ 自动清单: {auto_list.path} ({auto_list.material_count} 种物料)")
                print(f"✅ 汇总报告: {output}")
            return 0
    except Exception as exc:
        print(f"❌ 月初月末核验失败: {exc}")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())

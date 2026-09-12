from __future__ import annotations

import tempfile
import unittest
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo

import openpyxl

from pipeline.monthly_summary import (
    MailAttachmentCandidate,
    MonthlySource,
    aggregate_sources,
    build_report_frames,
    month_cutoff,
    parse_imap_internaldate,
    scan_month_files,
    select_latest_candidate,
)


TZ = ZoneInfo("Asia/Shanghai")


class MonthlySummaryTests(unittest.TestCase):
    def test_select_latest_candidate_respects_cutoff(self) -> None:
        candidates = [
            MailAttachmentCandidate("1", datetime(2026, 8, 31, 20, tzinfo=TZ), "物料情况和Excel文件", "总库存-a.xlsx", b"a"),
            MailAttachmentCandidate("2", datetime(2026, 8, 31, 23, 59, tzinfo=TZ), "物料情况和Excel文件", "总库存-b.xlsx", b"b"),
            MailAttachmentCandidate("3", datetime(2026, 9, 1, 0, 1, tzinfo=TZ), "物料情况和Excel文件", "总库存-c.xlsx", b"c"),
        ]
        chosen = select_latest_candidate(candidates, month_cutoff("2026-08"))
        self.assertEqual(chosen.uid, "2")

    def test_parse_imap_internaldate(self) -> None:
        parsed = parse_imap_internaldate(b'4 (INTERNALDATE "31-Aug-2026 23:59:00 +0800")')
        self.assertIsNotNone(parsed)
        self.assertEqual(parsed.isoformat(), "2026-08-31T23:59:00+08:00")

    def test_scan_auto_archive_takes_priority(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp) / "root"
            auto = root / "自动月末归档"
            root.mkdir()
            auto.mkdir()
            (root / "2502月底.xlsx").touch()
            (root / "2503月底.xlsx").touch()
            (auto / "2503月底.xlsx").touch()
            sources = scan_month_files(root, auto)
            self.assertEqual([source.month for source in sources], ["2025-02", "2025-03"])
            self.assertEqual(sources[1].source, "automatic")

    def test_read_and_aggregate_workbook(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            path = Path(temp) / "2502月底.xlsx"
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "库存表"
            ws.append([None, None, None, None])
            ws.append([None, None, None, None])
            ws.append([None, None, None, None])
            ws.append(["物料编码", "编号", "物料名称", "单位", "外仓入库总量", "外仓出库总量", "库存"])
            ws.append(["10403002000123", "00123", "测试物料", "公斤", 100, 60, 40])
            detail = wb.create_sheet("出入库明细表")
            detail.append([None, None])
            detail.append([None, None])
            detail.append(["录入日期", "库存变动类别"])
            detail.append(["2025/02/28 16:00:00", "出库"])
            wb.save(path)
            wb.close()

            result = aggregate_sources([MonthlySource("2025-02", path, "manual")])
            self.assertEqual(len(result.records), 1)
            self.assertEqual(result.records[0]["inbound"], 100)
            self.assertEqual(result.records[0]["outbound"], 60)
            self.assertEqual(result.quality[0]["month_end_rows"], 1)
            frames = build_report_frames(result)
            self.assertIn("物料总览", frames)
            self.assertEqual(frames["物料总览"].iloc[0]["累计净变化"], 40)


if __name__ == "__main__":
    unittest.main()

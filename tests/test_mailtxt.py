import importlib.util
import json
import tempfile
import unittest
from datetime import datetime
from pathlib import Path

import openpyxl


MODULE_PATH = Path(__file__).parents[1] / "script" / "050 mailtxt.py"
SPEC = importlib.util.spec_from_file_location("mailtxt", MODULE_PATH)
MAILTXT = importlib.util.module_from_spec(SPEC)
assert SPEC.loader is not None
SPEC.loader.exec_module(MAILTXT)


class MailDateHtmlTests(unittest.TestCase):
    def test_html_keeps_seconds_for_inventory_times(self):
        wb = openpyxl.Workbook()
        sheet = wb.active
        sheet["H3"] = datetime(2026, 9, 18, 8, 5, 6)
        sheet["M3"] = "2026-09-18T09:07:08+08:00"

        date, date2 = MAILTXT.get_dates(sheet)
        html = MAILTXT.construct_html_content(
            sheet, [], date, date2, 0, 0, 0, 0, 0, 0, 0
        )

        self.assertIn("2026-09-18 08:05:06", html)
        self.assertIn("2026-09-18 09:07:08", html)
        self.assertNotIn("2026-09-18</strong>", html)

    def test_html_prefers_source_mail_received_times(self):
        wb = openpyxl.Workbook()
        sheet = wb.active
        sheet["H3"] = "2026年09月18日"
        sheet["M3"] = "2026-09-18"
        with tempfile.TemporaryDirectory() as temp_dir:
            meta_path = Path(temp_dir) / "mail_meta.json"
            meta_path.write_text(json.dumps({
                "selected_heyu_da_received_at": "2026-09-18T10:11:12+08:00",
                "selected_waiting_received_at": "2026-09-18T13:14:15+08:00",
            }), encoding="utf-8")

            self.assertEqual(
                MAILTXT.get_dates(sheet, meta_path),
                ("2026-09-18 10:11:12", "2026-09-18 13:14:15"),
            )


if __name__ == "__main__":
    unittest.main()

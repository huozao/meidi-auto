import unittest

from pipeline.usage_trends import _companies_html, _month_cell, _rows_html


class UsageTrendHtmlTests(unittest.TestCase):
    months = ("2026-06", "2026-07", "2026-08")
    current_month = "2026-09"
    current_date = "2026-09-12"

    def rows(self):
        return [
            {
                "code": "MA001",
                "name": "物料一",
                "unit": "kg",
                "month": "2026-06",
                "outbound": 10,
                "date": "2026-06-30",
                "business_type": "领用出库",
                "company": "甲公司",
            },
            {
                "code": "MA001",
                "name": "物料一",
                "unit": "kg",
                "month": "2026-09",
                "outbound": 5,
                "date": self.current_date,
                "business_type": "领用出库",
                "company": "甲公司",
            },
        ]

    def test_zero_today_is_hidden_and_nonzero_is_inline(self):
        self.assertNotIn("今日 0", _month_cell(10, 0))
        self.assertIn("今日 +5", _month_cell(10, 5))

    def test_material_months_are_recent_first(self):
        rendered = _rows_html(self.rows(), self.months, self.current_month, self.current_date)
        self.assertLess(rendered.index("<th>2026-09</th>"), rendered.index("<th>2026-08</th>"))
        self.assertLess(rendered.index("<th>2026-08</th>"), rendered.index("<th>2026-07</th>"))
        self.assertNotIn("<th>今日变化</th>", rendered)
        self.assertNotIn("今日 0", rendered)
        self.assertIn("今日 +5", rendered)

    def test_company_months_are_recent_first(self):
        rows = self.rows()
        rendered = _companies_html(rows, rows, self.months, self.current_month, self.current_date)
        self.assertLess(rendered.index("<th>2026-09</th>"), rendered.index("<th>2026-08</th>"))
        self.assertLess(rendered.index("<th>2026-08</th>"), rendered.index("<th>2026-07</th>"))
        self.assertNotIn("<th>今日变化</th>", rendered)
        self.assertNotIn("今日 0", rendered)
        self.assertIn("今日 +5", rendered)


if __name__ == "__main__":
    unittest.main()

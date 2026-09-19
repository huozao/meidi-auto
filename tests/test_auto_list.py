import unittest

from pipeline.auto_list import build_auto_list_frame
from pipeline.monthly_summary import AggregationResult


class AutoListFormulaTests(unittest.TestCase):
    def test_each_warehouse_covers_two_weeks_and_sum_covers_month(self):
        records = [
            {"month": month, "short_code": "MA001", "code": "MA001", "outbound": 30}
            for month in ("2026-06", "2026-07", "2026-08")
        ]
        result = AggregationResult(records=records, quality=[], sources=[], movement_records=[])

        frame, months = build_auto_list_frame(result, "2026-08")

        self.assertEqual(months, ("2026-06", "2026-07", "2026-08"))
        self.assertEqual(frame.loc[0, "外应存(3月2周均)"], 15.0)
        self.assertEqual(frame.loc[0, "家应存(3月2周均)"], 15.0)
        self.assertEqual(frame.loc[0, "外应存(3月2周均)"] + frame.loc[0, "家应存(3月2周均)"], 30.0)


if __name__ == "__main__":
    unittest.main()

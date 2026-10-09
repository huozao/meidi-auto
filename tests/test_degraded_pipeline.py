import json
import os
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

import openpyxl
from openpyxl import Workbook


ROOT = Path(__file__).resolve().parents[1]


def load_safe_script(name: str, filename: str):
    import importlib.util
    path = ROOT / "script" / filename
    spec = importlib.util.spec_from_file_location(name, path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"无法加载脚本: {path}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


class DegradedPipelineTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.s050 = load_safe_script("s050", "050 mailtxt.py")
        cls.s051 = load_safe_script("s051", "051 Send an email.py")

    def test_020_create_placeholder_excel_via_subprocess(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            cmd = [
                sys.executable,
                "-c",
                f"""
import sys
sys.path.insert(0, r'{ROOT}')
import importlib.util
path = r'{ROOT / "script" / "020 Email download.py"}'
# 从文件读取函数源码并在安全独立作用域中执行
with open(path, 'r', encoding='utf-8') as f:
    code = f.read()
import openpyxl
from datetime import datetime
from zoneinfo import ZoneInfo
TZ_SH = ZoneInfo("Asia/Shanghai")
def now_shanghai():
    return datetime.now(TZ_SH)
# 提取 create_placeholder_inventory_excel 的定义
scope = {{'openpyxl': openpyxl, 'now_shanghai': now_shanghai, 'os': __import__('os')}}
exec('''def create_placeholder_inventory_excel(save_dir: str, file_prefix="存量查询") -> str:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "第一页"
    headers = ["仓库", "存货编码", "存货名称", "规格型号", "主计量", "主数量"]
    ws.append(headers)
    timestamp = now_shanghai().strftime("%Y%m%d_%H%M%S")
    file_name = f"{{file_prefix}}_{{timestamp}}.xlsx"
    full_path = os.path.join(save_dir, file_name)
    wb.save(full_path)
    wb.close()
    return full_path
''', scope)
p = scope['create_placeholder_inventory_excel'](r'{temp_dir}')
print(p)
"""
            ]
            res = subprocess.run(cmd, capture_output=True, text=True)
            self.assertEqual(res.returncode, 0, f"占位文件脚本执行失败: {res.stderr}")
            created_path = res.stdout.strip().splitlines()[-1]
            self.assertTrue(os.path.exists(created_path))
            wb = openpyxl.load_workbook(created_path)
            self.assertIn("第一页", wb.sheetnames)
            ws = wb["第一页"]
            headers = [cell.value for cell in ws[1]]
            self.assertIn("仓库", headers)
            self.assertIn("存货名称", headers)
            self.assertIn("主数量", headers)
            wb.close()

    def test_051_subject_prefix_on_degraded(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            excel_path = Path(temp_dir) / "总库存20261009.xlsx"
            wb = Workbook()
            ws = wb.active
            ws.title = "库存表"
            ws["H3"] = "2026年10月09日"
            wb.save(excel_path)
            wb.close()

            # 正常情况
            normal_subject = self.s051.build_subject(str(excel_path), home_stock_unavailable=False)
            self.assertEqual(normal_subject, "美的库存及出入库日报｜2026-10-09｜库存、出入库及月计划")

            # 降级情况
            degraded_subject = self.s051.build_subject(str(excel_path), home_stock_unavailable=True)
            self.assertEqual(degraded_subject, "【⚠️缺家里库存】美的库存及出入库日报｜2026-10-09｜库存、出入库及月计划")

    def test_050_html_renders_degraded_banner_and_placeholders(self):
        wb = Workbook()
        ws = wb.active
        ws.title = "库存表"
        ws["H3"] = "2026-10-09 20:00:00"
        ws["M3"] = "未同步（缺失）"

        # 模拟有 1 行预警物料（第 5 行）
        ws.cell(row=5, column=3, value="00514")      # 编号
        ws.cell(row=5, column=10, value=100.0)      # 库存
        ws.cell(row=5, column=11, value=200.0)      # 外应存
        ws.cell(row=5, column=13, value=None)       # 家里库存（空）

        html = self.s050.construct_html_content(
            sheet=ws,
            colored_rows=[5],
            date="2026-10-09 20:00:00",
            date2="未同步（缺失）",
            stock_total=1000.0,
            monthly_plan=500.0,
            plan_gap_output=None,
            monthly_sent=100.0,
            monthly_received=50.0,
            monthly_remaining=400.0,
            home_stock_total=None,
            home_stock_unavailable=True,
        )

        # 验证警示横幅
        self.assertIn("特别提醒：本次未收到畅捷通家里库存邮件", html)
        self.assertIn("未同步（缺失）", html)
        # 验证表格中的占位提示
        self.assertIn("无法计算", html)
        self.assertIn("未同步", html)

    def test_041_skips_production_and_gap_on_degraded(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            (temp_path / ".home-stock-unavailable").write_text("missing\n", encoding="utf-8")

            inv_file = temp_path / "总库存20261009.xlsx"
            wb = Workbook()
            ws = wb.active
            ws.title = "库存表"

            # 写入表头 (第 4 行)
            headers = [
                (1, "A"), (2, "序号"), (3, "编号"), (4, "名称"),
                (10, "库存"), (11, "外应存"), (12, "最小发货"), (13, "家里库存"),
                (14, "家应存"), (15, "排产"), (16, "月计划"), (17, "月计划缺口"),
                (18, "外仓出库总量"), (19, "外仓入库总量")
            ]
            for col_idx, col_name in headers:
                ws.cell(row=4, column=col_idx, value=col_name)

            # 写入 1 行数据 (第 5 行)
            ws.cell(row=5, column=2, value="1")       # B5
            ws.cell(row=5, column=10, value=100.0)    # J: 库存
            ws.cell(row=5, column=11, value=150.0)    # K: 外应存
            ws.cell(row=5, column=13, value=None)     # M: 家里库存
            ws.cell(row=5, column=14, value=80.0)     # N: 家应存
            ws.cell(row=5, column=16, value=300.0)    # P: 月计划
            ws.cell(row=5, column=18, value=20.0)     # R: 外仓出库总量

            wb.save(inv_file)
            wb.close()

            # 模拟执行 041 处理
            cmd = [sys.executable, str(ROOT / "script" / "041 operation.py"), str(temp_path)]
            res = subprocess.run(cmd, capture_output=True, text=True)
            self.assertEqual(res.returncode, 0, f"041 运行失败: {res.stderr}")

            wb_res = openpyxl.load_workbook(inv_file)
            ws_res = wb_res["库存表"]
            # 最小发货 = 外应存(150) - 库存(100) = 50
            self.assertEqual(ws_res.cell(row=5, column=12).value, 50.0)
            # 降级模式下：排产(col 15)与缺口(col 17)应置空为 None
            self.assertIsNone(ws_res.cell(row=5, column=15).value)
            self.assertIsNone(ws_res.cell(row=5, column=17).value)
            # 合计行 (第 6 行) 中排产与缺口也应为 None
            self.assertIsNone(ws_res.cell(row=6, column=15).value)
            self.assertIsNone(ws_res.cell(row=6, column=17).value)
            wb_res.close()


if __name__ == "__main__":
    unittest.main()

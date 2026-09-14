"""每日邮件中的物料和公司出库趋势 HTML。"""

from __future__ import annotations

import html
import os
import tempfile
from datetime import datetime
from pathlib import Path

import pandas as pd

from .archive import WebDavStore
from .monthly_summary import read_movement_workbook, scan_month_files


def _previous_month(value: str) -> str:
    year, month = map(int, value.split("-"))
    return f"{year - 1:04d}-12" if month == 1 else f"{year:04d}-{month - 1:02d}"


def _month_window(current_month: str) -> tuple[str, ...]:
    last = _previous_month(current_month)
    second = _previous_month(last)
    third = _previous_month(second)
    return (third, second, last)


def _fmt(value: object) -> str:
    try:
        number = float(value or 0)
    except (TypeError, ValueError):
        return "0"
    if number.is_integer():
        return f"{int(number):,}"
    return f"{number:,.2f}"


def _delta(value: object) -> str:
    try:
        number = float(value or 0)
    except (TypeError, ValueError):
        number = 0
    prefix = "+" if number > 0 else ""
    return f"{prefix}{_fmt(number)}"


def _short_code(value: object) -> str:
    digits = "".join(char for char in str(value or "") if char.isdigit())
    return digits[-5:].zfill(5) if digits else ""


def _display_unit(value: object) -> str:
    unit = str(value or "").strip()
    return "kg" if not unit or unit.lower() == "kg" or unit == "公斤" else unit


def _display_months(months: tuple[str, ...], current_month: str) -> tuple[str, ...]:
    """日报月份按最近到最远排列，最近月份放在第一列。"""
    return (current_month, *reversed(months))


def _month_cell(value: object, today: object = 0) -> str:
    """渲染月份单元格；仅在今日有非零变化时显示第二行提示。"""
    try:
        changed = float(today or 0) != 0
    except (TypeError, ValueError):
        changed = False
    extra = f"<div class='today-inline'>今日 {_delta(today)}</div>" if changed else ""
    return f"<div>{_fmt(value)}</div>{extra}"


def _current_file(data_dir: Path) -> Path:
    files = sorted(data_dir.glob("*总库存*.xlsx"), key=lambda path: path.stat().st_mtime)
    if not files:
        raise FileNotFoundError(f"未找到当前总库存文件: {data_dir}")
    return files[-1]


def _current_date_and_month(rows: list[dict], path: Path) -> tuple[str, str]:
    dates = [row.get("date", "") for row in rows if row.get("date")]
    if dates:
        current_date = max(dates)
        return current_date, current_date[:7]
    current = datetime.fromtimestamp(path.stat().st_mtime)
    return current.strftime("%Y-%m-%d"), current.strftime("%Y-%m")


def _load_history(root: Path | None, remote: bool) -> tuple[list[dict], tuple[str, ...]]:
    if root is None:
        return [], ()
    auto_dir = root / "自动月末归档"
    sources = scan_month_files(root, auto_dir)
    return [row for source in sources for row in read_movement_workbook(source.path, source.month)], tuple(source.month for source in sources)


def _download_remote_history(temp_root: Path) -> Path:
    base = os.getenv("NUTSTORE_WEBDAV_URL", "").strip()
    user = os.getenv("NUTSTORE_WEBDAV_USER", "").strip()
    password = os.getenv("NUTSTORE_WEBDAV_APP_PASSWORD", "").strip()
    remote_root = os.getenv("NUTSTORE_REMOTE_MONTHLY_ARCHIVE_DIR", "").strip()
    if not all((base, user, password, remote_root)):
        raise RuntimeError("未配置月度汇总目录或坚果云 WebDAV")
    store = WebDavStore(base, user, password)
    root = temp_root / "月度汇总"
    auto = root / "自动月末归档"
    auto.mkdir(parents=True)
    for name in store.list_files(f"{remote_root}/自动月末归档"):
        if name.endswith("月底.xlsx"):
            (auto / name).write_bytes(store.get(f"{remote_root}/自动月末归档/{name}"))
    return root


def _rows_html(rows: list[dict], months: tuple[str, ...], current_month: str, current_date: str) -> str:
    if not rows:
        return "<p>当前没有可用的出库明细。</p>"
    frame = pd.DataFrame(rows)
    frame = frame[frame["outbound"].map(lambda value: float(value or 0) > 0)].copy()
    if frame.empty:
        return "<p>近三个月和本月暂无出库记录。</p>"
    grouped = frame.groupby(["code", "name", "unit", "month"], as_index=False).agg(outbound=("outbound", "sum"))
    pivot = grouped.pivot_table(index=["code", "name", "unit"], columns="month", values="outbound", aggfunc="sum", fill_value=0).reset_index()
    for month in (*months, current_month):
        if month not in pivot:
            pivot[month] = 0
    pivot["_sort"] = pivot[list(months)].sum(axis=1)
    today = frame[frame["date"] == current_date].groupby("code")["outbound"].sum()
    pivot["今日变化"] = pivot["code"].map(today).fillna(0)
    pivot = pivot.sort_values(["_sort", current_month, "今日变化", "code"], ascending=[False, False, False, True])
    display_months = _display_months(months, current_month)
    headers = ["编号", *display_months, "单位"]
    out = ["<div class='table-wrap'><table class='trend'><colgroup>",
           "<col class='code-col'>", *["<col class='month-col'>" for _ in display_months],
           "<col class='unit-col'></colgroup><thead><tr>"]
    out += [f"<th>{html.escape(label)}</th>" for label in headers] + ["</tr></thead><tbody>"]
    for _, row in pivot.iterrows():
        cells = [f"<td class='code'>{_short_code(row['code'])}</td>"]
        cells.append(f"<td class='num current-month'>{_month_cell(row[current_month], row['今日变化'])}</td>")
        cells.extend(f"<td class='num'>{_fmt(row[month])}</td>" for month in reversed(months))
        cells.append(f"<td class='unit'>{html.escape(_display_unit(row['unit']))}</td>")
        out.append("<tr>" + "".join(cells) + "</tr>")
    total_cells = ["<td class='total-label'>全部物料出库合计</td>", f"<td class='num current-month'>{_month_cell(pivot[current_month].sum(), pivot['今日变化'].sum())}</td>"]
    total_cells.extend(f"<td class='num'>{_fmt(pivot[month].sum())}</td>" for month in reversed(months))
    total_cells.append("<td class='unit'>kg</td>")
    out.append("<tr class='total-row'>" + "".join(total_cells) + "</tr>")
    return "".join(out) + "</tbody></table></div>"


def _companies_html(rows: list[dict], all_rows: list[dict], months: tuple[str, ...], current_month: str, current_date: str) -> str:
    frame = pd.DataFrame(rows)
    if frame.empty:
        return "<p>没有可识别的领用公司出库记录。</p>"
    frame = frame[(frame["business_type"] == "领用出库") & (frame["outbound"].map(lambda value: float(value or 0) > 0))].copy()
    if frame.empty:
        return "<p>没有可识别的领用公司出库记录。</p>"
    history = frame[frame["month"].isin(months)].groupby("company")["outbound"].sum().sort_values(ascending=False)
    companies = list(history.head(5).index)
    if not companies:
        return "<p>近三个月没有领用公司出库记录。</p>"
    all_usage = pd.DataFrame(all_rows)
    all_usage = all_usage[(all_usage["business_type"] == "领用出库") & (all_usage["outbound"].map(lambda value: float(value or 0) > 0))].copy()
    display_months = _display_months(months, current_month)
    all_totals = {month: float(all_usage.loc[all_usage["month"] == month, "outbound"].sum()) for month in display_months}
    all_today_total = float(all_usage.loc[all_usage["date"] == current_date, "outbound"].sum())
    out = ["<div class='table-wrap'><table class='company'><colgroup><col class='rank-col'><col class='company-col'><col class='total-col'><col class='code-col'>",
           *["<col class='month-col'>" for _ in display_months],
           "<col class='unit-col'></colgroup><thead><tr><th>排名</th><th>公司</th><th>近3月总量</th><th>编号</th>",
           *[f"<th>{html.escape(month)}</th>" for month in display_months], "<th>单位</th></tr></thead><tbody>"]
    company_totals = {month: 0.0 for month in display_months}
    today_total = 0.0
    for rank, company in enumerate(companies, 1):
        company_rows = frame[frame["company"] == company]
        grouped = company_rows.groupby(["code", "name", "unit", "month"], as_index=False).agg(outbound=("outbound", "sum"))
        pivot = grouped.pivot_table(index=["code", "name", "unit"], columns="month", values="outbound", aggfunc="sum", fill_value=0).reset_index()
        for month in display_months:
            if month not in pivot:
                pivot[month] = 0
        pivot["_sort"] = pivot[list(months)].sum(axis=1)
        today = company_rows[company_rows["date"] == current_date].groupby("code")["outbound"].sum()
        pivot["今日变化"] = pivot["code"].map(today).fillna(0)
        for month in display_months:
            company_totals[month] += float(pivot[month].sum())
        today_total += float(pivot["今日变化"].sum())
        pivot = pivot.sort_values(["_sort", current_month, "code"], ascending=[False, False, True])
        rowspan = len(pivot)
        for index, (_, row) in enumerate(pivot.iterrows()):
            cells = []
            if index == 0:
                cells.extend([f"<td class='rank' rowspan='{rowspan}'>{rank}</td>", f"<td class='company-name' rowspan='{rowspan}'>{html.escape(str(company))}</td>", f"<td class='num company-total' rowspan='{rowspan}'>{_fmt(history[company])}</td>"])
            cells.append(f"<td class='code'>{_short_code(row['code'])}</td>")
            cells.append(f"<td class='num current-month'>{_month_cell(row[current_month], row['今日变化'])}</td>")
            cells.extend(f"<td class='num'>{_fmt(row[month])}</td>" for month in reversed(months))
            cells.append(f"<td class='unit'>{html.escape(_display_unit(row['unit']))}</td>")
            out.append("<tr>" + "".join(cells) + "</tr>")
    top_three_month_total = sum(company_totals[month] for month in months)
    all_three_month_total = sum(all_totals[month] for month in months)
    total_cells = ["<td class='total-label' colspan='2'>前五大公司合计</td>", f"<td class='num'>{_fmt(top_three_month_total)}</td>", "<td></td>", f"<td class='num current-month'>{_month_cell(company_totals[current_month], today_total)}</td>"]
    total_cells.extend(f"<td class='num'>{_fmt(company_totals[month])}</td>" for month in reversed(months))
    total_cells.append("<td class='unit'>kg</td>")
    out.append("<tr class='total-row'>" + "".join(total_cells) + "</tr>")
    all_cells = ["<td class='total-label' colspan='2'>全部领用出库合计</td>", f"<td class='num'>{_fmt(all_three_month_total)}</td>", "<td></td>", f"<td class='num current-month'>{_month_cell(all_totals[current_month], all_today_total)}</td>"]
    all_cells.extend(f"<td class='num'>{_fmt(all_totals[month])}</td>" for month in reversed(months))
    all_cells.append("<td class='unit'>kg</td>")
    out.append("<tr class='total-row'>" + "".join(all_cells) + "</tr>")
    concentration_cells = ["<td class='total-label' colspan='4'>前五大公司领用集中度</td>"]
    current_concentration = f"{(company_totals[current_month] / all_totals[current_month] * 100):.1f}%" if all_totals[current_month] else "—"
    today_concentration = f"{(today_total / all_today_total * 100):.1f}%" if today_total and all_today_total else ""
    today_concentration_cell = f"<div class='today-inline'>{today_concentration} 今日</div>" if today_concentration else ""
    concentration_cells.append(f"<td class='num current-month'><div>{current_concentration}</div>{today_concentration_cell}</td>")
    concentration_cells.extend(
        f"<td class='num'>{(company_totals[month] / all_totals[month] * 100):.1f}%</td>" if all_totals[month] else "<td class='num'>—</td>"
        for month in reversed(months)
    )
    concentration_cells.append("<td class='unit'>%</td>")
    out.append("<tr class='total-row concentration-row'>" + "".join(concentration_cells) + "</tr>")
    return "".join(out) + "</tbody></table></div>"


def build_usage_sections(data_dir: Path, monthly_root: Path | None = None) -> str:
    current_path = _current_file(data_dir)
    # 当前文件中的明细不要求是月末文件，直接用于本月和今日统计。
    current_rows = read_movement_workbook(current_path, datetime.fromtimestamp(current_path.stat().st_mtime).strftime("%Y-%m"))
    current_date, current_month = _current_date_and_month(current_rows, current_path)
    # 文件可能从邮件/坚果云复制后拥有新的 mtime，月份以明细中的业务日期为准。
    for row in current_rows:
        row["month"] = current_month
    months = _month_window(current_month)
    history_rows: list[dict] = []
    with tempfile.TemporaryDirectory(prefix="meidi-usage-") as temp:
        root = monthly_root
        if root is None and os.getenv("NUTSTORE_REMOTE_MONTHLY_ARCHIVE_DIR", "").strip():
            try:
                root = _download_remote_history(Path(temp))
            except Exception as exc:
                return f"<h2>出库趋势</h2><p class='warn'>历史月度汇总读取失败：{html.escape(str(exc))}</p>"
        try:
            history_rows, _ = _load_history(root, False) if root else ([], ())
        except Exception as exc:
            return f"<h2>出库趋势</h2><p class='warn'>历史月度汇总读取失败：{html.escape(str(exc))}</p>"
    rows = [row for row in history_rows if row.get("month") in months] + current_rows
    recognized = [row for row in rows if row.get("company") != "未备注"]
    return ("<h2>物料出库趋势</h2>"
            + _rows_html(rows, months, current_month, current_date)
            + "<h2>前五大领用公司及物料明细</h2>"
            + _companies_html(recognized, rows, months, current_month, current_date))

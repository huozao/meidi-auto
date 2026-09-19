#!/usr/bin/env python3
"""将现有旧口径自动清单安全迁移为两周均值并发布到 WebDAV。

该入口只处理自动清单，不下载邮箱、不生成日报、不发送邮件，也不改写月度报告。
只有明确识别到旧版“3月周均”表头时才会把外应存、家应存数值乘 2；
新口径或未知表头均停止，避免重复放大或误改人工清单。
"""

from __future__ import annotations

import hashlib
import io
import os
import sys
from pathlib import Path

import openpyxl

REPO_ROOT = Path(__file__).resolve().parent.parent
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from pipeline.archive import WebDavStore


OLD_EXTERNAL_HEADERS = ("外应存(1周)", "外应存(3月周均)")
OLD_HOME_HEADERS = ("家应存(1周)", "家应存(3月周均)")
NEW_EXTERNAL = "外应存(3月2周均)"
NEW_HOME = "家应存(3月2周均)"


def _number(value: object) -> int | float | None:
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return value
    return None


def migrate_legacy_list(source: bytes) -> tuple[bytes, int]:
    """迁移一份旧清单，返回新文件字节和被调整的物料行数。"""
    input_stream = io.BytesIO(source)
    wb = openpyxl.load_workbook(input_stream)
    try:
        ws = wb.active
        headers = {str(cell.value).strip(): cell.column for cell in ws[1] if cell.value is not None}
        ext_col = next((headers.get(label) for label in OLD_EXTERNAL_HEADERS if headers.get(label)), None)
        home_col = next((headers.get(label) for label in OLD_HOME_HEADERS if headers.get(label)), None)
        if ext_col and home_col:
            ws.cell(row=1, column=ext_col, value=NEW_EXTERNAL)
            ws.cell(row=1, column=home_col, value=NEW_HOME)
            changed = 0
            for row in range(2, ws.max_row + 1):
                ext = _number(ws.cell(row=row, column=ext_col).value)
                home = _number(ws.cell(row=row, column=home_col).value)
                if ext is None and home is None:
                    continue
                if ext is not None:
                    ws.cell(row=row, column=ext_col, value=round(ext * 2, 2))
                if home is not None:
                    ws.cell(row=row, column=home_col, value=round(home * 2, 2))
                changed += 1
            output = io.BytesIO()
            wb.save(output)
            return output.getvalue(), changed

        if headers.get(NEW_EXTERNAL) and headers.get(NEW_HOME):
            raise ValueError("自动清单已经是 3月2周均口径，无需重复迁移")
        discovered = ", ".join(headers) or "(空)"
        raise ValueError(
            f"无法安全识别自动清单口径，实际表头: {discovered}；"
            f"需同时包含旧版“外应存(1周/3月周均)”和“家应存(1周/3月周均)”"
        )
    finally:
        wb.close()


def _required(name: str) -> str:
    value = os.getenv(name, "").strip()
    if not value:
        raise RuntimeError(f"缺少环境变量: {name}")
    return value


def publish() -> int:
    base = _required("NUTSTORE_WEBDAV_URL")
    user = _required("NUTSTORE_WEBDAV_USER")
    password = _required("NUTSTORE_WEBDAV_APP_PASSWORD")
    remote_path = _required("NUTSTORE_REMOTE_AUTO_LIST_FILE")
    store = WebDavStore(base, user, password)
    original = store.get(remote_path)
    updated, changed = migrate_legacy_list(original)
    store.put(remote_path, updated)
    print(
        f"✅ 已发布两周均值自动清单: {remote_path} | 调整 {changed} 行 | "
        f"SHA256 {hashlib.sha256(original).hexdigest()[:12]} -> {hashlib.sha256(updated).hexdigest()[:12]}"
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(publish())

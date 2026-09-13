from __future__ import annotations

"""STEP CARD: 将每日邮件附件及展示文件备份到 YYYYMM 归档目录。"""

import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from pipeline.archive import archive_daily_outputs


def main(argv: list[str] | None = None) -> int:
    argv = argv or sys.argv
    data_dir = Path(argv[1] if len(argv) >= 2 else "data").resolve()
    archive_daily_outputs(data_dir)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

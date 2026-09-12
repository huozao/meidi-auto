"""仓库流水线统一入口。

按固定顺序执行生产步骤，并共享同一个数据目录。
支持本地运行与 GitHub Actions 运行。
"""

from __future__ import annotations

import argparse
import glob
import importlib.util
import json
import os
import subprocess
import sys
import time
from pathlib import Path

from dotenv import load_dotenv

from pipeline.models import PipelineStep
from pipeline.steps import CLEANUP_STEP, PRODUCTION_STEPS
from pipeline.validators import missing_step_files, validate_step_inputs, validate_step_output


REQUIRED_ENV_KEYS: tuple[str, ...] = (
    "EMAIL_ADDRESS_QQ",
    "EMAIL_PASSWORD_QQ|EMAIL_PASSWOR_QQ",
    "RECIPIENT_EMAILS",
)

RETRYABLE_STEPS: tuple[str, ...] = ("020 Email download.py",)
DEFAULT_RETRY_COUNT = 2
DEFAULT_RETRY_BACKOFF_SECONDS = 2
DEFAULT_IN_PROCESS_STEPS: tuple[str, ...] = ("050 image.py", "050 mailtxt.py", "051 Send an email.py")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="运行 MeidiAuto 自动化流水线")
    parser.add_argument(
        "--data-dir",
        default="data",
        help="所有步骤共用的数据输出目录（默认：./data）",
    )
    parser.add_argument(
        "--script-dir",
        default="script",
        help="步骤脚本目录（默认：./script）",
    )
    parser.add_argument(
        "--stop-on-error",
        action="store_true",
        help="任一步骤失败后立即停止",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="只打印执行计划，不实际运行",
    )
    parser.add_argument(
        "--check",
        action="store_true",
        help="仅检查环境变量与脚本文件后退出",
    )
    parser.add_argument(
        "--report-file",
        default="",
        help="可选：JSON 报告输出路径（相对路径以仓库根目录为基准）",
    )
    parser.add_argument(
        "--clean-only",
        action="store_true",
        help="仅运行清理脚本（010 clean.py）后退出",
    )
    parser.add_argument(
        "--clean-after-run",
        action="store_true",
        help="流水线结束后执行清理脚本（用于删除测试/冗余产物）",
    )
    parser.add_argument(
        "--list-steps",
        action="store_true",
        help="仅列出可运行步骤后退出",
    )
    parser.add_argument(
        "--only-step",
        action="append",
        default=[],
        help="仅运行指定步骤（可重复；可填完整文件名、编号前缀或关键字，如 030 / '041 operation.py'）",
    )
    parser.add_argument(
        "--retry-count",
        type=int,
        default=DEFAULT_RETRY_COUNT,
        help="失败重试次数（仅对网络敏感步骤生效，默认：2）",
    )
    parser.add_argument(
        "--retry-backoff",
        type=int,
        default=DEFAULT_RETRY_BACKOFF_SECONDS,
        help="重试退避秒数基值（默认：2；第n次等待 base*n 秒）",
    )
    parser.add_argument(
        "--in-process-steps",
        default=",".join(DEFAULT_IN_PROCESS_STEPS),
        help="逗号分隔：以进程内函数方式执行的步骤列表（默认包含 050/051）",
    )
    return parser.parse_args()


def _has_env_group(value: str) -> bool:
    return any(os.getenv(k.strip()) for k in value.split("|"))


def check_environment() -> list[str]:
    missing: list[str] = []
    for key in REQUIRED_ENV_KEYS:
        if not _has_env_group(key):
            missing.append(key)
    return missing


def resolve_steps(only_step_args: list[str], all_steps: tuple[PipelineStep, ...]) -> tuple[tuple[PipelineStep, ...], list[str]]:
    if not only_step_args:
        return all_steps, []

    requests: list[str] = []
    for raw in only_step_args:
        for part in raw.split(","):
            token = part.strip()
            if token:
                requests.append(token)

    resolved: list[PipelineStep] = []
    invalid: list[str] = []
    for token in requests:
        token_lower = token.lower()
        matched = [
            step
            for step in all_steps
            if step.filename.lower() == token_lower
            or step.filename.lower().startswith(token_lower)
            or token_lower in step.filename.lower()
        ]
        if not matched:
            invalid.append(token)
            continue
        for step in matched:
            if step not in resolved:
                resolved.append(step)

    return tuple(resolved), invalid


def _load_step_module(script_path: Path):
    module_name = f"step_{script_path.stem.replace(" ", "_")}_{abs(hash(script_path))}"
    spec = importlib.util.spec_from_file_location(module_name, script_path)
    if spec is None or spec.loader is None:
        return None, "无法创建模块加载器"
    module = importlib.util.module_from_spec(spec)
    try:
        spec.loader.exec_module(module)
    except Exception as exc:
        return None, f"模块加载失败: {exc}"
    return module, ""


def _run_step_in_process(script_path: Path, data_dir: Path) -> tuple[int, str]:
    module, load_msg = _load_step_module(script_path)
    if module is None:
        return 1, load_msg
    if not hasattr(module, "main"):
        return 127, "缺少 main() 入口"

    main_func = getattr(module, "main")
    argv = [str(script_path), str(data_dir)]
    old_argv = sys.argv
    sys.argv = argv
    try:
        # 优先尝试 main(argv)，失败再回退 main()
        try:
            result = main_func(argv)
        except TypeError:
            result = main_func()
    finally:
        sys.argv = old_argv

    if result is None:
        return 0, ""
    if isinstance(result, int):
        return result, ""
    return 0, ""


def run_step(step: PipelineStep, script_dir: Path, data_dir: Path, retry_count: int = DEFAULT_RETRY_COUNT, retry_backoff: int = DEFAULT_RETRY_BACKOFF_SECONDS, in_process_steps: frozenset[str] = frozenset()) -> tuple[bool, float]:
    script_path = script_dir / step.filename
    if not script_path.exists():
        msg = f"❌ 缺少步骤脚本：{script_path}"
        if step.required:
            print(msg)
            return False, 0.0
        print(f"⚠️ {msg}（可选步骤，已跳过）")
        return True, 0.0

    input_ok, input_msg = validate_step_inputs(step, data_dir)
    if not input_ok:
        print(f"❌ {step.filename} 输入校验失败：{input_msg}")
        return False, 0.0

    cmd = [sys.executable, str(script_path), str(data_dir)]
    attempts = 1 + (max(0, retry_count) if step.filename in RETRYABLE_STEPS else 0)
    total_elapsed = 0.0

    for attempt in range(1, attempts + 1):
        print(f"\n🚀 正在运行：{' '.join(cmd)}（尝试 {attempt}/{attempts}）")
        started = time.time()
        if step.filename in in_process_steps:
            return_code, in_process_msg = _run_step_in_process(script_path, data_dir)
            if in_process_msg:
                print(f"⚠️ 进程内执行提示: {in_process_msg}")
            elapsed = time.time() - started
            total_elapsed += elapsed
            if return_code == 0:
                completed_returncode = 0
            else:
                completed_returncode = return_code
        else:
            env = os.environ.copy()
            env.setdefault("PYTHONIOENCODING", "utf-8")
            env.setdefault("PYTHONUTF8", "1")
            completed = subprocess.run(cmd, capture_output=True, text=False, env=env)
            elapsed = time.time() - started
            total_elapsed += elapsed

            stdout_text = _decode_subprocess_output(completed.stdout)
            stderr_text = _decode_subprocess_output(completed.stderr)

            if stdout_text:
                print(stdout_text)
            if stderr_text:
                print(stderr_text)
            completed_returncode = completed.returncode

        if completed_returncode == 0:
            output_ok, output_msg = validate_step_output(step, data_dir)
            if not output_ok:
                print(f"❌ {step.filename} 输出校验失败：{output_msg}")
                return False, total_elapsed
            print(f"✅ {step.filename} 执行完成，用时 {total_elapsed:.2f}s")
            return True, total_elapsed

        print(f"❌ {step.filename} 执行失败（退出码={completed_returncode}），本次用时 {elapsed:.2f}s")
        if attempt < attempts:
            wait_seconds = max(1, retry_backoff * attempt)
            print(f"🔁 {step.filename} 将在 {wait_seconds}s 后重试...")
            time.sleep(wait_seconds)

    return False, total_elapsed


def _decode_subprocess_output(raw: bytes) -> str:
    """尽量兼容 Windows/跨平台编码，避免因 gbk/utf-8 不一致导致解码异常。"""
    if not raw:
        return ""
    for encoding in ("utf-8", "gbk"):
        try:
            return raw.decode(encoding)
        except UnicodeDecodeError:
            continue
    return raw.decode("utf-8", errors="replace")


def write_report(report_file: str, payload: dict, root: Path) -> None:
    if not report_file:
        return
    path = Path(report_file)
    if not path.is_absolute():
        path = (root / path).resolve()
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"🧾 运行报告已写入：{path}")


def main() -> int:
    # 本地 WSL 运行读取仓库根目录 .env；GitHub Actions 仍由环境变量注入。
    load_dotenv(Path(__file__).resolve().with_name(".env"))
    args = parse_args()
    in_process_steps = frozenset([item.strip() for item in args.in_process_steps.split(",") if item.strip()])

    if args.clean_only and args.clean_after_run:
        print("❌ 参数冲突：--clean-only 与 --clean-after-run 不能同时使用。")
        return 2

    root = Path(__file__).resolve().parent
    script_dir = (root / args.script_dir).resolve()
    data_dir = (root / args.data_dir).resolve()

    print(f"📁 脚本目录：{script_dir}")
    print(f"📁 数据目录：{data_dir}")
    if args.list_steps:
        print("📋 可运行步骤：")
        for index, step in enumerate(PRODUCTION_STEPS, start=1):
            print(f"{index:02d}. {step.filename}")
        print("🧹 清理步骤：010 clean.py（通过 --clean-only 或 --clean-after-run 启用）")
        return 0

    selected_steps, invalid_steps = resolve_steps(args.only_step, PRODUCTION_STEPS)
    if invalid_steps:
        print(f"❌ 未识别的步骤标识：{', '.join(invalid_steps)}")
        print("可先执行 `python main.py --list-steps` 查看可用步骤。")
        return 2
    if not selected_steps and not args.clean_only:
        print("❌ 没有可执行步骤。可先执行 `python main.py --list-steps`。")
        return 2

    print(f"📋 步骤总数：{len(selected_steps)}")

    if not script_dir.exists():
        print(f"❌ 脚本目录不存在：{script_dir}")
        return 2

    steps_to_check: tuple[PipelineStep, ...] = selected_steps
    if args.clean_only or args.clean_after_run:
        steps_to_check = selected_steps + (CLEANUP_STEP,)
    missing_files = missing_step_files(script_dir, steps_to_check)
    missing_env = check_environment()

    if args.dry_run:
        planned_steps: list[str] = []
        if args.clean_only:
            planned_steps.append(CLEANUP_STEP.filename)
        else:
            planned_steps.extend(step.filename for step in selected_steps)
            if args.clean_after_run:
                planned_steps.append(CLEANUP_STEP.filename)
        for index, filename in enumerate(planned_steps, start=1):
            print(f"{index:02d}. {filename}")
        if args.clean_only:
            print("🧹 当前仅执行清理步骤。")
        elif args.clean_after_run:
            print("🧹 将在流水线结束后追加清理步骤。")
        print("🧪 演练模式执行完成。")
        return 0

    if args.check:
        if missing_files:
            print(f"❌ 缺少必需步骤脚本：{', '.join(missing_files)}")
        else:
            print("✅ 必需步骤脚本齐全。")

        if missing_env:
            print(f"⚠️ 缺少环境变量：{', '.join(missing_env)}")
        else:
            print("✅ 必需环境变量已就绪。")

        check_details = {
            "step_contracts": [
                {
                    "filename": step.filename,
                    "required": step.required,
                    "input_patterns": list(step.input_patterns),
                    "output_patterns": list(step.output_patterns),
                    "script_exists": (script_dir / step.filename).exists(),
                }
                for step in steps_to_check
            ],
            "missing_step_files": missing_files,
            "missing_env": missing_env,
            "retry": {
                "retryable_steps": list(RETRYABLE_STEPS),
                "retry_count": args.retry_count,
                "retry_backoff": args.retry_backoff,
            },
        }
        write_report(
            args.report_file,
            {
                "mode": "check",
                "success": not missing_files,
                "details": check_details,
            },
            root,
        )

        return 0 if not missing_files else 1

    if args.clean_only:
        ok, _ = run_step(CLEANUP_STEP, script_dir, data_dir, args.retry_count, args.retry_backoff, in_process_steps)
        if ok:
            print("🧹 清理任务执行完成。")
            return 0
        print("❌ 清理任务执行失败。")
        return 1

    os.makedirs(data_dir, exist_ok=True)

    failures: list[str] = []
    total_time = 0.0
    started_at = time.strftime("%Y-%m-%d %H:%M:%S")

    for step in selected_steps:
        ok, elapsed = run_step(step, script_dir, data_dir, args.retry_count, args.retry_backoff, in_process_steps)
        total_time += elapsed
        if not ok:
            failures.append(step.filename)
            if args.stop_on_error:
                break

    if args.clean_after_run:
        print("\n🧹 开始执行收尾清理（010 clean.py）...")
        clean_ok, clean_elapsed = run_step(CLEANUP_STEP, script_dir, data_dir, args.retry_count, args.retry_backoff, in_process_steps)
        total_time += clean_elapsed
        if not clean_ok:
            failures.append(CLEANUP_STEP.filename)

    print("\n================ 流水线执行汇总 ================")
    success = not failures
    if failures:
        print(f"❌ 失败步骤（{len(failures)}）：{', '.join(failures)}")
        print(f"⏱️ 总耗时：{total_time:.2f}s")
    else:
        print("🎉 所有步骤执行成功")
        print(f"⏱️ 总耗时：{total_time:.2f}s")

    write_report(
        args.report_file,
        {
            "started_at": started_at,
            "elapsed_seconds": round(total_time, 2),
            "success": success,
            "failed_steps": failures,
            "step_count": len(selected_steps),
            "data_dir": str(data_dir),
            "script_dir": str(script_dir),
        },
        root,
    )

    return 0 if success else 1


if __name__ == "__main__":
    raise SystemExit(main())

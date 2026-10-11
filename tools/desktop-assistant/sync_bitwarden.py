#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Bitwarden Secrets Manager 环境变量同步工具 (BWS Sync)
功能：
1. 优先使用 Bitwarden 官方 Secrets Manager (bws CLI)；
2. 自动定位系统或本目录下的 bws / bws.exe；
3. 从环境变量或本地 .bws_token 读取机器账户令牌；
4. 秒级拉取最新环境变量并写入本地 .env，全程 100% 物理隔离，不触碰任何个人生活密码。
"""

import os
import sys
import shutil
import subprocess
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
TOKEN_FILE = BASE_DIR / ".bws_token"
ENV_FILE = BASE_DIR / ".env"


def find_bws_executable() -> str:
    """查找可用的 bws 可执行文件"""
    is_windows = sys.platform == "win32"

    if is_windows:
        # 1. Windows: 优先找当前目录下的 bws.exe
        local_bws_win = BASE_DIR / "bws.exe"
        if local_bws_win.exists():
            return str(local_bws_win)
        # 2. Windows: PATH
        path_bws_exe = shutil.which("bws.exe")
        if path_bws_exe:
            return path_bws_exe
        path_bws = shutil.which("bws")
        if path_bws:
            return path_bws
    else:
        # 1. Linux/macOS: 优先找当前目录下的 bws (二进制)
        local_bws_nix = BASE_DIR / "bws"
        if local_bws_nix.exists() and os.access(str(local_bws_nix), os.X_OK):
            return str(local_bws_nix)
        # 2. Linux/macOS: PATH
        path_bws = shutil.which("bws")
        if path_bws:
            return path_bws

    return ""


def get_access_token() -> str:
    """获取 BWS_ACCESS_TOKEN"""
    # 1. 环境变量优先
    token = os.getenv("BWS_ACCESS_TOKEN", "").strip()
    if token:
        return token

    # 2. 读取本地 .bws_token 文件
    if TOKEN_FILE.exists():
        try:
            return TOKEN_FILE.read_text(encoding="utf-8").strip()
        except Exception:
            pass

    return ""


def save_access_token(token: str):
    """保存 Token 到本地 .bws_token"""
    TOKEN_FILE.write_text(token.strip(), encoding="utf-8")


def sync_from_bws(log_func=print) -> bool:
    """核心同步方法"""
    bws_bin = find_bws_executable()
    if not bws_bin:
        log_func("❌ 未找到 bws 命令行工具。请确保已下载 bws / bws.exe。")
        return False

    token = get_access_token()
    if not token:
        log_func("❌ 未找到 BWS_ACCESS_TOKEN。请先设置环境变量或在 .bws_token 中配置。")
        return False

    log_func("🔒 正在连接 Bitwarden Secrets Manager 拉取云端凭据...")
    env = os.environ.copy()
    env["BWS_ACCESS_TOKEN"] = token

    try:
        # 运行 bws secret list -o env
        cmd = [bws_bin, "secret", "list", "-o", "env"]
        # Windows 上若通过 pythonw 运行，隐藏黑框
        creationflags = 0
        if sys.platform == "win32":
            creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0x08000000)

        res = subprocess.run(
            cmd,
            env=env,
            capture_output=True,
            text=True,
            encoding="utf-8",
            creationflags=creationflags,
            timeout=15
        )

        if res.returncode != 0:
            err = (res.stderr or res.stdout).strip()
            log_func(f"❌ 拉取凭据失败: {err}")
            return False

        content = res.stdout.strip()
        if not content:
            log_func("⚠️ 云端未配置任何密钥，未更新本地 .env。")
            return False

        # 写入 .env 文件
        ENV_FILE.write_text(content + "\n", encoding="utf-8")
        log_func(f"✅ 凭据同步成功！已写入本地配置文件: {ENV_FILE.name}")
        return True

    except subprocess.TimeoutExpired:
        log_func("❌ 连接 Bitwarden 超时，请检查网络连接。")
        return False
    except Exception as e:
        log_func(f"❌ 同步异常: {e}")
def update_bws_secrets(updates: dict, log_func=print) -> bool:
    """将本地修改的变量反向推送到 Bitwarden Secrets Manager 云端"""
    bws_bin = find_bws_executable()
    token = get_access_token()
    if not bws_bin or not token:
        log_func("⚠️ 未配置 bws 工具或令牌，无法反向同步至云端。")
        return False

    env = os.environ.copy()
    env["BWS_ACCESS_TOKEN"] = token
    creationflags = 0
    if sys.platform == "win32":
        creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0x08000000)

    try:
        import json
        # 1. 查询现有的 secrets 映射
        res = subprocess.run(
            [bws_bin, "secret", "list", "-o", "json"],
            env=env,
            capture_output=True,
            text=True,
            encoding="utf-8",
            creationflags=creationflags,
            timeout=15
        )
        if res.returncode != 0:
            log_func(f"⚠️ 查询云端机密列表失败: {res.stderr or res.stdout}")
            return False

        secrets = json.loads(res.stdout)
        key_to_id = {s["key"]: s["id"] for s in secrets if "key" in s and "id" in s}

        updated_count = 0
        for k, v in updates.items():
            if k in key_to_id:
                secret_id = key_to_id[k]
                edit_res = subprocess.run(
                    [bws_bin, "secret", "edit", "--value", str(v), secret_id],
                    env=env,
                    capture_output=True,
                    text=True,
                    encoding="utf-8",
                    creationflags=creationflags,
                    timeout=15
                )
                if edit_res.returncode == 0:
                    updated_count += 1
                    log_func(f"☁️ 【{k}】已成功同步至 Bitwarden 云端！")
                else:
                    log_func(f"⚠️ 更新【{k}】到云端失败: {edit_res.stderr or edit_res.stdout}")

        return updated_count > 0

    except Exception as e:
        log_func(f"⚠️ 反向同步至云端异常: {e}")
        return False


if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "--set-token":
        if len(sys.argv) > 2:
            save_access_token(sys.argv[2])
            print("✅ 令牌已成功保存到 .bws_token！")
        else:
            print("用法: python sync_bitwarden.py --set-token <TOKEN>")
        sys.exit(0)

    success = sync_from_bws()
    sys.exit(0 if success else 1)

# -*- coding: utf-8 -*-
"""
美的仓储自动化控制台 (Meidi Auto Assistant)
整合功能：
1. 手动下载最新物料情况邮件与附件（存至“美的发货”目录）
2. 触发 GitHub Actions 每日流水线（自动运行 run-daily.yml）
3. 一键打开附件目录与配置管理
支持美观原生图形界面 (GUI) 与控制台双模式。
"""

import os
import sys
import json
import time
import email
import imaplib
import threading
import urllib.request
import urllib.error
from datetime import datetime
from email.header import decode_header
from pathlib import Path

# GUI 组件预先安全导入
try:
    import tkinter as tk
    from tkinter import ttk, messagebox, scrolledtext
    TKINTER_AVAILABLE = True
except Exception:
    TKINTER_AVAILABLE = False

# ================================
# 基础配置与文件路径
# ================================
BASE_DIR = Path(__file__).resolve().parent
ENV_FILE = BASE_DIR / ".env"
CONFIG_FILE = BASE_DIR / "config.json"

# 默认配置
# 默认配置模版（敏感凭据留空，支持从 .env、项目根目录或界面配置）
DEFAULT_CONFIG = {
    "IMAP_SERVER": "imap.qq.com",
    "IMAP_PORT": 993,
    "QQ_EMAIL": "",
    "QQ_AUTH_CODE": "",
    "KEYWORD": "物料情况",
    "SAVE_DIR": "./downloads",
    "GITHUB_REPO": "huozao/meidi-auto",
    "GITHUB_WORKFLOW": "run-daily.yml",
    "GITHUB_REF": "main",
    "GITHUB_TOKEN": "",
}


def load_env_or_config() -> dict:
    """加载配置：优先读取 config.json，其次 .env，最后默认值。"""
    config = dict(DEFAULT_CONFIG)

    # 尝试读取 .env
    if ENV_FILE.exists():
        try:
            with open(ENV_FILE, "r", encoding="utf-8") as f:
                for line in f:
                    line = line.strip()
                    if line and not line.startswith("#") and "=" in line:
                        k, v = line.split("=", 1)
                        config[k.strip()] = v.strip().strip("\"'")
        except Exception:
            pass

    # 尝试读取 config.json
    if CONFIG_FILE.exists():
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                saved = json.load(f)
                config.update(saved)
        except Exception:
            pass

    return config


def save_config(config: dict):
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(config, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f"保存配置失败: {e}")


# ================================
# 核心业务 1：下载邮件与附件
# ================================
def download_mail_and_attachments(log_callback=print) -> bool:
    config = load_env_or_config()
    server = config.get("IMAP_SERVER", "imap.qq.com")
    port = int(config.get("IMAP_PORT", 993))
    email_user = config.get("QQ_EMAIL", "")
    email_pwd = config.get("QQ_AUTH_CODE", "")
    keyword = config.get("KEYWORD", "物料情况")
    save_dir = config.get("SAVE_DIR", r"C:\Users\ishel\Desktop\美的发货")

    if not email_user or not email_pwd:
        log_callback("❌ 错误：未配置邮箱账号或授权码！")
        return False

    os.makedirs(save_dir, exist_ok=True)
    log_callback(f"📂 附件保存路径: {save_dir}")
    log_callback(f"📡 正在连接邮箱服务器: {server}:{port}...")

    mail = None
    try:
        mail = imaplib.IMAP4_SSL(server, port)
        mail.login(email_user, email_pwd)
        log_callback("✅ 邮箱登录成功，正在检索收件箱...")
        mail.select("inbox")

        status, data = mail.search(None, "ALL")
        if status != "OK" or not data or not data[0]:
            log_callback("⚠️ 收件箱为空或查询失败。")
            return False

        all_ids = data[0].split()
        recent_ids = all_ids[-35:]  # 最近 35 封
        log_callback(f"📨 正在分析最近 {len(recent_ids)} 封邮件...")

        matched_msg = None
        matched_subject = ""
        for msg_id in reversed(recent_ids):
            status, msg_data = mail.fetch(msg_id, "(RFC822)")
            if status != "OK" or not msg_data or not msg_data[0]:
                continue
            raw_email = msg_data[0][1]
            msg = email.message_from_bytes(raw_email)

            subject_raw = msg.get("Subject", "")
            decoded_parts = decode_header(subject_raw)
            subject = ""
            for p, enc in decoded_parts:
                if isinstance(p, bytes):
                    subject += p.decode(enc or "utf-8", errors="ignore")
                else:
                    subject += str(p)

            if keyword in subject:
                matched_msg = msg
                matched_subject = subject
                log_callback(f"🎯 命中目标邮件: {subject}")
                break

        if not matched_msg:
            log_callback(f"⚠️ 未在最近邮件中找到包含关键词【{keyword}】的邮件。")
            return False

        # 解析时间
        date_str = matched_msg.get("Date", "")
        time_str = "latest"
        month_str = datetime.now().strftime("%Y%m")
        date_tuple = email.utils.parsedate_tz(date_str) if date_str else None
        if date_tuple:
            recv_time = datetime.fromtimestamp(email.utils.mktime_tz(date_tuple))
            time_str = recv_time.strftime("%Y%m%d_%H%M")
            month_str = recv_time.strftime("%Y%m")
        log_callback(f"🕒 邮件时间戳: {time_str} (对应归档月份: {month_str})")

        # 智能匹配/自动创建月份归档文件夹 (如 202610)
        final_save_dir = os.path.join(save_dir, month_str)
        if not os.path.exists(final_save_dir):
            os.makedirs(final_save_dir, exist_ok=True)
            log_callback(f"📁 自动创建月份归档文件夹: {final_save_dir}")
        else:
            log_callback(f"📁 已匹配现存月份归档文件夹: {final_save_dir}")

        # 提取 HTML 正文
        html_content = None
        for part in matched_msg.walk():
            if part.get_content_type() == "text/html" and not part.get_filename():
                charset = part.get_content_charset() or "utf-8"
                html_content = part.get_payload(decode=True).decode(charset, errors="replace")
                break

        if html_content:
            safe_subj = "".join(c for c in matched_subject if c.isalnum() or c in " _-").strip()
            html_name = f"{safe_subj}_{time_str}.html"
            html_path = os.path.join(final_save_dir, html_name)
            with open(html_path, "w", encoding="utf-8") as f:
                f.write(html_content)
            log_callback(f"📄 正文已保存: {html_name}")

        # 提取附件
        att_count = 0
        for part in matched_msg.walk():
            if part.get_content_maintype() == "multipart":
                continue
            if part.get("Content-Disposition") is None:
                continue

            raw_fname = part.get_filename()
            if raw_fname:
                parts = decode_header(raw_fname)
                fname = ""
                for p, enc in parts:
                    if isinstance(p, bytes):
                        fname += p.decode(enc or "utf-8", errors="ignore")
                    else:
                        fname += str(p)

                base, ext = os.path.splitext(fname)
                save_fname = f"{base}_{time_str}{ext}"
                file_dest = os.path.join(final_save_dir, save_fname)
                with open(file_dest, "wb") as f:
                    f.write(part.get_payload(decode=True))
                log_callback(f"📎 附件已下载: {save_fname}")
                att_count += 1

        if att_count == 0:
            log_callback("ℹ️ 该邮件未包含附件或未识别到附件。")

        log_callback(f"🎉 下载完成！共保存 {att_count} 个附件至【{final_save_dir}】。")
        return True

    except Exception as e:
        log_callback(f"❌ 下载过程发生异常: {e}")
        return False
    finally:
        if mail:
            try:
                mail.logout()
            except Exception:
                pass


# ================================
# 核心业务 2：触发 GitHub Actions
# ================================
def trigger_github_action(log_callback=print) -> bool:
    config = load_env_or_config()
    repo = config.get("GITHUB_REPO", "huozao/meidi-auto")
    workflow = config.get("GITHUB_WORKFLOW", "run-daily.yml")
    ref = config.get("GITHUB_REF", "main")
    token = config.get("GITHUB_TOKEN", "").strip()

    if not token or token == "请在此处粘贴您的Token":
        log_callback("❌ 错误：尚未配置 GITHUB_TOKEN，请先配置有效 Token！")
        return False

    url = f"https://api.github.com/repos/{repo}/actions/workflows/{workflow}/dispatches"
    log_callback(f"🚀 正在向 GitHub 请求触发流水线...")
    log_callback(f"🎯 目标仓库: {repo} | 工作流: {workflow} | 分支: {ref}")

    payload = json.dumps({"ref": ref}).encode("utf-8")
    req = urllib.request.Request(url, data=payload, method="POST")
    req.add_header("Authorization", f"Bearer {token}")
    req.add_header("Accept", "application/vnd.github.v3+json")
    req.add_header("Content-Type", "application/json")
    req.add_header("User-Agent", "MeidiAuto-Assistant/1.0")

    try:
        with urllib.request.urlopen(req) as resp:
            status_code = resp.status
            log_callback(f"📡 GitHub 响应状态码: {status_code}")
            if status_code == 204:
                log_callback("✅ 成功！GitHub Actions 流水线已被成功唤醒并开始运行！")
                log_callback("💡 您可以稍后在邮箱查收处理结果。")
                return True
            else:
                log_callback(f"⚠️ 响应状态码非 204: {status_code}")
                return False
    except urllib.error.HTTPError as e:
        err_msg = e.read().decode("utf-8", errors="ignore")
        log_callback(f"❌ 触发失败 (HTTP {e.code}): {e.reason}")
        if e.code == 401:
            log_callback("   [原因排查] Token 无效或已过期，请核对 Token。")
        elif e.code == 403:
            log_callback("   [原因排查] Token 权限不足，请确认是否勾选了 repo 和 workflow 权限。")
        elif e.code == 404:
            log_callback("   [原因排查] 仓库或工作流文件不存在，请检查仓库配置。")
        if err_msg:
            log_callback(f"   详细错误: {err_msg}")
        return False
    except Exception as e:
        log_callback(f"❌ 请求发生网络或连接异常: {e}")
        return False


# ================================
# 现代化 GUI 界面 (Tkinter)
# ================================
class AssistantGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("美的仓储自动化控制台")

        # 毫秒级直接居中定位，消除 update_idletasks() 带来的闪烁与卡顿
        w, h = 640, 560
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        x = max(0, (sw - w) // 2)
        y = max(0, (sh - h) // 2)
        self.root.geometry(f"{w}x{h}+{x}+{y}")
        self.root.minsize(580, 480)

        # 窗口配色
        self.bg_color = "#f4f6f9"
        self.card_bg = "#ffffff"
        self.text_color = "#2d3748"
        self.accent_color = "#3182ce"
        self.root.configure(bg=self.bg_color)

        # 挂载应用精简专属图标
        icon_path = BASE_DIR / "app_icon.ico"
        if icon_path.exists():
            try:
                self.root.iconbitmap(str(icon_path))
            except Exception:
                pass

        self.config = load_env_or_config()
        self.create_widgets()

    def create_widgets(self):
        # 顶部标题栏卡片
        header_frame = tk.Frame(self.root, bg="#2b6cb0", height=64)
        header_frame.pack(fill="x", side="top")

        title_label = tk.Label(
            header_frame,
            text="📦 美的仓储自动化控制台",
            font=("Microsoft YaHei", 15, "bold"),
            bg="#2b6cb0",
            fg="#ffffff"
        )
        title_label.pack(side="left", padx=20, pady=12)

        subtitle_label = tk.Label(
            header_frame,
            text="一键下载邮件 · 触发每日对账 · 极速快捷",
            font=("Microsoft YaHei", 9),
            bg="#2b6cb0",
            fg="#bee3f8"
        )
        subtitle_label.pack(side="right", padx=20, pady=16)

        # 中部操作卡片区域
        action_frame = tk.Frame(self.root, bg=self.bg_color, pady=12, padx=16)
        action_frame.pack(fill="x")

        # 按钮网格布局
        btn_font = ("Microsoft YaHei", 10, "bold")

        btn1 = tk.Button(
            action_frame,
            text="📥 手动下载最新邮件附件\n(拉取物料情况至美的发货)",
            font=btn_font,
            bg="#ebf8ff",
            fg="#2b6cb0",
            activebackground="#bee3f8",
            relief="groove",
            bd=1,
            cursor="hand2",
            padx=10,
            pady=8,
            command=self.start_download_thread
        )
        btn1.grid(row=0, column=0, padx=8, pady=6, sticky="nsew")

        btn2 = tk.Button(
            action_frame,
            text="🚀 立即触发 GitHub 流水线\n(启动美的美对账发送)",
            font=btn_font,
            bg="#f0fff4",
            fg="#276749",
            activebackground="#c6f6d5",
            relief="groove",
            bd=1,
            cursor="hand2",
            padx=10,
            pady=8,
            command=self.start_trigger_thread
        )
        btn2.grid(row=0, column=1, padx=8, pady=6, sticky="nsew")

        btn3 = tk.Button(
            action_frame,
            text="📂 打开当月存放文件夹\n(自动定位最新月份归档)",
            font=("Microsoft YaHei", 9),
            bg="#edf2f7",
            fg="#4a5568",
            relief="groove",
            bd=1,
            cursor="hand2",
            pady=3,
            command=self.open_save_dir
        )
        btn3.grid(row=1, column=0, padx=8, pady=4, sticky="nsew")

        btn4 = tk.Button(
            action_frame,
            text="⚙️ 配置与使用说明",
            font=("Microsoft YaHei", 9),
            bg="#edf2f7",
            fg="#4a5568",
            relief="groove",
            bd=1,
            cursor="hand2",
            pady=5,
            command=self.show_config_dialog
        )
        btn4.grid(row=1, column=1, padx=8, pady=4, sticky="nsew")

        action_frame.columnconfigure(0, weight=1)
        action_frame.columnconfigure(1, weight=1)

        # 下方控制台日志区域
        log_container = tk.Frame(self.root, bg=self.bg_color, padx=16, pady=4)
        log_container.pack(fill="both", expand=True)

        log_head = tk.Frame(log_container, bg=self.bg_color)
        log_head.pack(fill="x", pady=(0, 4))

        tk.Label(
            log_head,
            text="🖥️ 执行状态与实时日志",
            font=("Microsoft YaHei", 9, "bold"),
            bg=self.bg_color,
            fg="#4a5568"
        ).pack(side="left")

        clear_btn = tk.Button(
            log_head,
            text="清空日志",
            font=("Microsoft YaHei", 8),
            bg="#e2e8f0",
            fg="#4a5568",
            relief="flat",
            padx=6,
            command=self.clear_logs
        )
        clear_btn.pack(side="right")

        self.log_text = scrolledtext.ScrolledText(
            log_container,
            wrap="word",
            bg="#1a202c",
            fg="#e2e8f0",
            font=("Consolas", 9),
            insertbackground="white",
            relief="sunken",
            bd=1
        )
        self.log_text.pack(fill="both", expand=True)

        # 底部状态栏
        status_bar = tk.Frame(self.root, bg="#edf2f7", height=24)
        status_bar.pack(fill="x", side="bottom")

        token_hint = self.config.get("GITHUB_TOKEN", "")
        masked_token = f"ghp_...{token_hint[-4:]}" if len(token_hint) > 8 else "未配置"
        self.status_lbl = tk.Label(
            status_bar,
            text=f"状态: 就绪 | 仓库: {self.config.get('GITHUB_REPO')} | Token: {masked_token}",
            font=("Microsoft YaHei", 8),
            bg="#edf2f7",
            fg="#718096"
        )
        self.status_lbl.pack(side="left", padx=10, pady=2)

        self.log("💡 控制台就绪。点击上方按钮可直接执行对应任务。")

    def log(self, message: str):
        self.log_text.configure(state="normal")
        ts = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert("end", f"[{ts}] {message}\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def clear_logs(self):
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")

    def start_download_thread(self):
        self.log("----------------------------------------")
        self.log("🚀 开始执行：手动下载邮件任务...")
        t = threading.Thread(target=self._run_download, daemon=True)
        t.start()

    def _run_download(self):
        ok = download_mail_and_attachments(log_callback=self.log)
        self.log("----------------------------------------")

    def start_trigger_thread(self):
        self.log("----------------------------------------")
        self.log("🚀 开始执行：触发 GitHub Actions 流水线...")
        t = threading.Thread(target=self._run_trigger, daemon=True)
        t.start()

    def _run_trigger(self):
        ok = trigger_github_action(log_callback=self.log)
        self.log("----------------------------------------")

    def open_save_dir(self):
        def _do_open():
            base_dir = self.config.get("SAVE_DIR", r"C:\Users\ishel\Desktop\美的发货")
            # 优先定位当月归档子目录 (如 202610)
            cur_month = datetime.now().strftime("%Y%m")
            target_dir = os.path.join(base_dir, cur_month)
            if not os.path.exists(target_dir):
                target_dir = base_dir

            if os.path.exists(target_dir):
                try:
                    # Windows 原生毫秒级启动，完全不经 cmd.exe，绝不卡死界面
                    if hasattr(os, "startfile"):
                        os.startfile(target_dir)
                    else:
                        import subprocess
                        subprocess.Popen(["explorer.exe", target_dir])
                    self.log(f"📂 已为您秒开文件夹: {target_dir}")
                except Exception as e:
                    self.log(f"⚠️ 打开文件夹异常: {e}")
            else:
                self.log(f"⚠️ 目标文件夹不存在: {target_dir}")

        threading.Thread(target=_do_open, daemon=True).start()

    def show_config_dialog(self):
        diag = tk.Toplevel(self.root)
        diag.title("⚙️ 配置与说明")
        diag.geometry("520x420")
        diag.transient(self.root)
        diag.grab_set()

        # 居中子窗口
        dx = self.root.winfo_x() + 50
        dy = self.root.winfo_y() + 50
        diag.geometry(f"+{dx}+{dy}")

        tk.Label(
            diag,
            text="快捷控制台说明与配置修改",
            font=("Microsoft YaHei", 11, "bold"),
            fg="#2b6cb0"
        ).pack(anchor="w", padx=20, pady=(15, 5))

        desc = (
            "1. 【手动下载邮件】：从 QQ 邮箱自动检索关键词“物料情况”最新邮件，"
            "将附件与 HTML 正文保存至【美的发货】目录。\n"
            "2. 【触发流水线】：向 GitHub API 发送指令触发 Run MeidiAuto Pipeline，"
            "即使家里库存缺失也会发送应急通报。\n"
            "3. 如需更新 Token，在下方修改后点击【保存配置】即可。"
        )
        tk.Label(
            diag,
            text=desc,
            font=("Microsoft YaHei", 8),
            justify="left",
            fg="#4a5568",
            wraplength=480
        ).pack(anchor="w", padx=20, pady=5)

        form = tk.Frame(diag, padx=20, pady=10)
        form.pack(fill="x")

        # GitHub Token
        tk.Label(form, text="GitHub Token (永久以 ghp_ 开头):", font=("Microsoft YaHei", 9)).grid(row=0, column=0, sticky="w", pady=4)
        token_entry = tk.Entry(form, font=("Consolas", 9), width=45)
        token_entry.insert(0, self.config.get("GITHUB_TOKEN", ""))
        token_entry.grid(row=1, column=0, sticky="w", pady=(0, 8))

        # 保存目录
        tk.Label(form, text="附件保存目录:", font=("Microsoft YaHei", 9)).grid(row=2, column=0, sticky="w", pady=4)
        dir_entry = tk.Entry(form, font=("Microsoft YaHei", 9), width=45)
        dir_entry.insert(0, self.config.get("SAVE_DIR", ""))
        dir_entry.grid(row=3, column=0, sticky="w", pady=(0, 8))

        # 邮箱账号
        tk.Label(form, text="QQ 邮箱账号:", font=("Microsoft YaHei", 9)).grid(row=4, column=0, sticky="w", pady=4)
        email_entry = tk.Entry(form, font=("Microsoft YaHei", 9), width=45)
        email_entry.insert(0, self.config.get("QQ_EMAIL", ""))
        email_entry.grid(row=5, column=0, sticky="w", pady=(0, 8))

        def save_and_close():
            self.config["GITHUB_TOKEN"] = token_entry.get().strip()
            self.config["SAVE_DIR"] = dir_entry.get().strip()
            self.config["QQ_EMAIL"] = email_entry.get().strip()
            save_config(self.config)
            messagebox.showinfo("成功", "配置已成功保存！", parent=diag)
            self.status_lbl.configure(text=f"状态: 配置已更新 | 仓库: {self.config.get('GITHUB_REPO')}")
            diag.destroy()

        btn_box = tk.Frame(diag, pady=10)
        btn_box.pack(fill="x", padx=20)
        tk.Button(
            btn_box,
            text="💾 保存配置",
            font=("Microsoft YaHei", 9, "bold"),
            bg="#3182ce",
            fg="white",
            padx=14,
            pady=4,
            command=save_and_close
        ).pack(side="right")


# ================================
# 控制台 CLI 模式 (备用与兜底)
# ================================
def run_cli_menu():
    while True:
        os.system("cls" if os.name == "nt" else "clear")
        print("=" * 56)
        print("         📦 美的仓储自动化控制台 (CLI 模式)")
        print("=" * 56)
        print("  [1] 📥 手动下载最新邮件附件 (物料情况 -> 美的发货)")
        print("  [2] 🚀 触发 GitHub Actions 每日流水线")
        print("  [3] 📂 打开当月存放文件夹 (自动定位最新月份归档)")
        print("  [4] ⚙️ 查看当前配置说明")
        print("  [0] ❌ 退出")
        print("=" * 56)
        choice = input("请选择操作序号 [0-4]: ").strip()

        if choice == "1":
            print("\n正在下载邮件...")
            download_mail_and_attachments(print)
            input("\n按回车键返回菜单...")
        elif choice == "2":
            print("\n正在触发 GitHub 流水线...")
            trigger_github_action(print)
            input("\n按回车键返回菜单...")
        elif choice == "3":
            cfg = load_env_or_config()
            base_dir = cfg.get("SAVE_DIR", r"C:\Users\ishel\Desktop\美的发货")
            cur_month = datetime.now().strftime("%Y%m")
            target_dir = os.path.join(base_dir, cur_month)
            if not os.path.exists(target_dir):
                target_dir = base_dir

            if os.path.exists(target_dir):
                if hasattr(os, "startfile"):
                    os.startfile(target_dir)
                else:
                    import subprocess
                    subprocess.Popen(["explorer.exe", target_dir])
                print(f"已秒开文件夹: {target_dir}")
            else:
                print(f"目录不存在: {target_dir}")
            input("\n按回车键返回菜单...")
        elif choice == "4":
            cfg = load_env_or_config()
            print("\n【当前配置】")
            print(f"  邮箱账号: {cfg.get('QQ_EMAIL')}")
            print(f"  保存目录: {cfg.get('SAVE_DIR')}")
            print(f"  目标仓库: {cfg.get('GITHUB_REPO')} ({cfg.get('GITHUB_WORKFLOW')})")
            t = cfg.get('GITHUB_TOKEN', '')
            print(f"  Token:    {'ghp_...' + t[-4:] if len(t) > 8 else '未配置'}")
            input("\n按回车键返回菜单...")
        elif choice in ("0", "q", "exit"):
            print("退出控制台。")
            break


# ================================
# 主入口：优先启动 GUI，失败退回 CLI
# ================================
def main():
    # 若显式传入 --cli 或环境无 Tkinter 则运行命令行菜单
    if "--cli" in sys.argv or not TKINTER_AVAILABLE:
        run_cli_menu()
        return

    try:
        root = tk.Tk()
        app = AssistantGUI(root)
        root.mainloop()
    except Exception as e:
        print(f"无法启动图形界面 ({e})，自动切换为控制台模式：")
        run_cli_menu()


if __name__ == "__main__":
    main()

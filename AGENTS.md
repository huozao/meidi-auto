# meidi-auto

美的存量对帐流水线。**每天在生产运行**，由 [`gmail-watcher`](../gmail-watcher/) 触发：
Gmail 收到「骏都对帐表」→ Cloud Run 命中关键词 → dispatch 本仓 `run-daily.yml`
→ 下载邮件、合并 Excel、着色计算、生成图与正文、发信给三个收件人。

先读 `README.md`（含运行方式、必需环境变量、历史工具的去向）和
`docs/PIPELINE_FLOW.md`（主流程一页图）。

## 动手前必须知道的

- **改动前先看今天跑没跑**。`gh run list --repo huozao/meidi-auto --limit 3`。
  正常每天一次，由 Gmail 事件驱动，不是固定 cron。

- **`Actions 绿灯 ≠ 业务跑通`**。`run-daily.yml` 有一条「仅 `020 Email download.py`
  失败时记为软失败并放行」的分支，命中它时 run 依然是 success。判据要落在
  artifact 里 `run-report.json` 的 `failed_steps`，以及日志里有没有
  `✅ 邮件发送成功`——不是 run 的 conclusion。

- **收件人在 `RECIPIENT_EMAILS` 这个 secret 里，不在代码里**。2026-09-08 收编
  新架构时才从硬编码改过来，生产值是三个地址。⚠️ 历史上 `MeidiAuto2` 那份
  `.env` 里只有一个地址（开发自测用），照抄会**静默少发两个人**，不报任何错。

- **密码 secret 的名字是 `EMAIL_PASSWOR_QQ`**（历史误拼，少一个 D）。代码
  `os.getenv("EMAIL_PASSWORD_QQ") or os.getenv("EMAIL_PASSWOR_QQ")` 两个都读，
  改 secret 名之前先确认两边都改。

- **凭据正本在 `infra/secrets/meidi-auto.enc.env`**（SOPS）。改那里不等于改生产——
  GitHub Secrets 没有读回接口，必须 `gh secret set` 推回才生效。

## 环境变量与凭据流转（AI 必读）

本仓具有两条清晰分立的凭据管理体系：

1. **生产流水线凭据（GitHub Actions）**：
   - 凭据正本在 `infra/secrets/meidi-auto.enc.env`（SOPS 加密）；
   - 生产消费方为 GitHub Secrets（`EMAIL_ADDRESS_QQ`, `EMAIL_PASSWOR_QQ`, `RECIPIENT_EMAILS`）；
   - 改动 SOPS 后需通过 `gh secret set` 推送生产生效。

2. **桌面助手与本地开发凭据（Desktop Assistant / Local Dev）**：
   - 凭据正本托管在 **Bitwarden Secrets Manager (BWS)**（组织：`DevOps-Workspac`，项目：`meidi-auto`）；
   - **AI 获取凭据指引**：若本地或 Windows 桌面环境缺失 `.env`，**严禁要求用户手动粘贴明文凭据**。直接调用 `python tools/desktop-assistant/sync_bitwarden.py`（依赖本地已被 gitignore 保护的 `.bws_token` 与官方 `bws`/`bws.exe` 工具），0.3 秒内静默生成 `.env`；
   - **双向同步机制**：控制台（`meidi_console.py` / `meidi_console.pyw`）修改配置并保存时，会自动异步调用 `bws secret edit` 回写云端；手机/网页端修改后，桌面端点击【🔄 同步 Bitwarden 凭据】即可拉取最新；
   - **安全与隔离底线**：使用的服务账号为 AI 专用 Machine Account，可访问本项目与通用账号本 `accounts`（同一令牌也在 `~/.config/bws/.bws_token`），**物理级无法触碰用户的个人密码库**。注意：同步脚本只写 `tools/desktop-assistant/.env`；仓库根 `.env` 是主程序配置，不由 BWS 生成。严禁将 `.env`、`.bws_token`、`bws` 二进制提交至 Git。

## 提交

走 PR（`gh pr create --repo huozao/meidi-auto`）。合并到 main 不会自动部署，
本仓没有部署概念——下次被 dispatch 时就跑新代码。

想验证改动，用 `gh workflow run run-daily.yml --repo huozao/meidi-auto --ref <分支>`
在分支上实跑一次；**它会真的发邮件**，验完记得说明。

工作区级规矩见顶层 [`AGENTS.md`](../AGENTS.md)。

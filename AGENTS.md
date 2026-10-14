# meidi-auto

美的存量对帐流水线：下载邮件、处理 Excel、生成日报、发送邮件并归档。
本文件适用于独立 clone；父工作区存在时再读其跨仓规则，缺少工作区不阻断本地任务。

## 先读与代码入口

- [README.md](README.md)：安装、无业务副作用的检查、执行边界。
- [docs/CURRENT_HANDOFF.md](docs/CURRENT_HANDOFF.md)：已确认口径、带日期的运行证据与待验证项。
- [docs/PIPELINE_FLOW.md](docs/PIPELINE_FLOW.md)：主流程；步骤与输入输出契约在 `pipeline/steps.py`、`pipeline/validators.py`。
- `main.py` 负责参数与编排，`script/` 执行业务，`tools/monthly_summary.py` 承载独立月度工具。
- 月度口径见 [docs/MONTHLY_SUMMARY_SPEC.md](docs/MONTHLY_SUMMARY_SPEC.md)，本地调试见 [docs/development.md](docs/development.md)。

## 业务与授权边界

- `main.py` 默认完整运行、`--only-step` 和手动 dispatch 都可能真实读邮箱、发信或写 WebDAV；独立数据目录不隔离真实账户。真实联调、发信、归档与清理必须在用户明确授权范围内。
- 影响线上运行或准备真实联调时，先只读检查近期 Actions，避免重复处理；纯本地文档/单元测试无需访问 GitHub 或生产。
- 日报由 `run-daily.yml` 的 `workflow_dispatch` 触发。Actions success 不能证明业务完成：IMAP 不可达可跳过，仅 020 失败可软失败放行。检查报告的 `success`、`skipped`、`failed_steps`、邮件发送日志；用户确认收到强于发送成功日志。
- `RECIPIENT_EMAILS` 来自配置，不凭历史收件人数推断现值；密码兼容 `EMAIL_PASSWORD_QQ` / `EMAIL_PASSWOR_QQ`。不得擅改收件人或凭据。
- 邮件、附件、Excel/HTML、月末快照、同步文件、`.env` 与日志不入 Git。工作区部署的凭据管理按其 SOPS 规则，独立 clone 使用自行保管的本地配置。
- 未经明确授权不 commit、push、merge、创建 PR、修改 Secrets 或其他远程资源。获授权交付时走 PR；main 更新后下一次工作流会执行新代码。文档里的交付流程不等于授权。

## 最小验证与维护

从仓库根目录使用安装了 `requirements.txt` 的 Python：

```bash
python main.py --list-steps
python main.py --dry-run
python main.py --check
python tools/generate_module_catalog.py --check
git diff --check
```

`--check` 缺少环境变量只告警，返回 0 不能证明凭据完整或邮箱可用。
改代码追加受影响测试（现有入口：`python -m unittest discover -s tests -v`）。
纯文档改动验证链接、命令即可，不要求完整实跑。

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
   - **安全与隔离底线**：使用的服务账号为 Machine Account，仅拥有 `meidi-auto` 项目权限，**物理级无法触碰用户的个人密码库**（100% 隔离）。严禁将 `.env`、`.bws_token`、`bws` 二进制提交至 Git。

## 提交

业务行为、步骤、输入输出或验收方式变化时，同批更新受影响的现有契约、STEP CARD
与导航；新功能按需要补最小文档，无影响不制造记录。历史证据标明日期，
待验证不得写成线上事实，旧评分/迁移讨论不覆盖当前源码。

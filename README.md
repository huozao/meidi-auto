# meidi-auto

美的存量对帐流水线：下载邮件附件、合并和计算 Excel、生成日报、发送邮件并归档。
每日入口是 `main.py`；每日工作流通过 `workflow_dispatch` 由 Gmail 监听器或人工触发。
独立月度核验与汇总不加入每日步骤。

## 从这里开始

| 需要做什么 | 入口 |
|---|---|
| 查看业务口径、最近取证与未验证项 | [当前交接](docs/CURRENT_HANDOFF.md)（运行结果均有日期，不能代替当前探测） |
| 理解步骤与输入输出 | [主流程](docs/PIPELINE_FLOW.md)、`pipeline/steps.py`、`pipeline/validators.py` |
| 月末快照、汇总和自动清单 | [月度契约](docs/MONTHLY_SUMMARY_SPEC.md) |
| 本地调试、报告、模块表维护 | [开发帮助](docs/development.md) |
| AI 修改规则与授权边界 | [AGENTS.md](AGENTS.md) |
| 查询旧教材、评分、迁移讨论 | [历史 README](docs/migrations/2026-09-13-readme-history.md)（不作为当前操作说明） |

## 安装与本地检查

在仓库根目录使用 Python 3.12 和独立虚拟环境，安装依赖：

```bash
python -m pip install -r requirements.txt
python main.py --list-steps
python main.py --dry-run
```

## 必需环境变量

- `EMAIL_ADDRESS_QQ`
- `EMAIL_PASSWORD_QQ`（兼容历史变量 `EMAIL_PASSWOR_QQ`）
- `IMAP_SERVER`（可选，默认 `imap.qq.com`）

### 凭据获取与同步说明

- **CI/CD 生产环境**：由 GitHub Actions 自动注入生产密钥。正本维护在 `infra/secrets/meidi-auto.enc.env`（SOPS 加密）。
- **桌面助手与本地开发**：统一托管于 **Bitwarden Secrets Manager**（项目：`meidi-auto`）。本地执行 `python tools/desktop-assistant/sync_bitwarden.py` 即可秒级拉取生成 `.env`，桌面控制台支持双向热重载。

## 月末邮件与物料汇总

月度汇总已迁移到 WSL 中的 `tools/monthly_summary.py`，设计和数据边界见
[`docs/MONTHLY_SUMMARY_SPEC.md`](docs/MONTHLY_SUMMARY_SPEC.md)。它不加入每日
`main.py` 生产步骤，默认按指定月份选择最后一天 24:00 前最新的
“物料情况和Excel文件”邮件，保存月末快照，再扫描历史月末文件并输出物料分析 Excel。
工作簿是否为合格月末表，以 `库存表` 表头上方是否包含目标月份最后一天（如
`2026年08月31日`）为主判据；当天没有出入库流水只作辅助提示。
报告还会根据 `出入库明细表` 的“备注”统一公司名称，输出公司月度/年度领用物料、
借用归还、退货换货等业务明细，并在 `图表` 工作表中生成公司和业务类型趋势图。
备注为空时，`MA1111`按“重庆工厂”统计，`MA1141`和`MS1121`按子库编号作为公司名统计。

本地配置放在仓库根目录 `.env`（不提交），至少包括 `EMAIL_ADDRESS_QQ`、
`EMAIL_PASSWORD_QQ`、`MONTHLY_ARCHIVE_DIR`。常用命令：

```bash
# 只用现有月末文件生成报告，不访问邮箱
python tools/monthly_summary.py --aggregate-only

# 下载指定月份月末邮件、归档并生成全量报告
python tools/monthly_summary.py --month 2026-08

# 只检查月份、字段和数据质量，不写入归档或报告
python tools/monthly_summary.py --aggregate-only --dry-run

# 下载最新目标邮件供人工查看
python tools/download_latest_mail.py
```

## 历史工具的去向与替代方案

2026-09-08 从 MeidiAuto 收编本架构时，下列脚本随主线一起清理。它们都还在 git 历史里，用
`git log --all --oneline -- '<路径>'` 可以取回。

| 已移除 | 原用途 | 现在怎么做 |
|---|---|---|
| `script/sync_env_to_github.py` | 用 PyGithub 把 `.env` 批量推成 GitHub Secrets | `gh secret set -f .env --repo <owner>/<repo>`，CLI 原生支持 env 文件批量导入，省掉 PyGithub 依赖 |
| `chek.py`、`多网站.py`、`script/main local.py` | 早期试验/本地入口 | 已被 `main.py` 的 `--check` / `--dry-run` / `--only-step` 取代 |
| `script/050 image local.py`、`script/052 send email.py` | 早期本地版分支 | 已被 `050 image.py` / `051 Send an email.py` 取代 |
| `script/月汇总.py`、`script/月汇总绘图.py` | 月度汇总与绘图 | 从不在流水线内的本地手工工具，且硬编码了本机桌面路径；如需恢复，取回后必须先改掉绝对路径 |
| `script/统一格式.py`、`script/企业消息整理.py` | 一次性数据整理 | 同上，按需从历史取回 |
| `.github/workflows/run_script.yml` | 早期手动触发工作流 | 已被 `run-daily.yml` 取代 |

## `.env` 安全与联调建议

不建议把真实 `.env` 文件直接发给任何人（包括我），避免密码与邮箱凭据泄露。

你截图这个页面（`Settings -> Secrets and variables -> Actions`）**不是上传 `.env` 文件本身**，而是把 `.env` 里的每个键值拆开，分别配置成 GitHub Secrets（这是正确位置）。

推荐做法：

1. 本地创建 `.env` 并自行保管（加入 `.gitignore`，不要提交到仓库）。
2. 我可以基于你提供的**脱敏样例**帮你检查格式是否正确，例如：

```env
EMAIL_ADDRESS_QQ=demo@example.com
EMAIL_PASSWORD_QQ=********
RECIPIENT_EMAILS=a@example.com,b@example.com
IMAP_SERVER=imap.qq.com
```

3. 本地验证时先跑：

```bash
python main.py --check
python tools/generate_module_catalog.py --check
```

上述检查不执行邮箱下载、Excel 处理、发信或 WebDAV 写入。`--dry-run` 仅列计划；
`--check` 检查脚本与环境变量，缺少环境变量只告警，退出码 0 不代表配置完整或凭据有效。
如加 `--report-file`，检查命令会写本地报告。代码改动的现有测试入口：

```bash
python -m unittest discover -s tests -v
git diff --check
```

## 真实执行边界

`python main.py` 默认执行完整流水线，会访问邮箱、发送邮件，并按配置归档到本地或 WebDAV。
`--only-step` 也会真正执行选中步骤；独立 `--data-dir` 不隔离账户和远程写入。
人工 dispatch、月度工具和清理也有真实副作用，应在明确授权且确认目标、收件人与输入后执行。

获得真实联调授权后，可使用以下完整运行方式：

```bash
python main.py --data-dir data-local --stop-on-error --report-file data-local/run-report.json
```

执行前先查看近期 Actions，避免与自动任务重复发信。不要把完整运行作为文档修改的默认验证。
`--clean-only` / `--clean-after-run` 会删除产物，使用前须检查目标范围。

## 配置与数据

本地 `.env` 由 `main.py` 加载；Actions 从 Secrets/Variables 注入配置。真实凭据、
附件、Excel/HTML、快照、报告、同步文件和日志不提交 Git。

| 配置 | 用途 |
|---|---|
| `EMAIL_ADDRESS_QQ` | 邮箱账号 |
| `EMAIL_PASSWORD_QQ` / `EMAIL_PASSWOR_QQ` | 授权码，兼容历史误拼 |
| `RECIPIENT_EMAILS` | 必需，逗号分隔收件人；执行前核对实际配置 |
| `IMAP_SERVER` | IMAP 主机，工作流空值回退 `imap.qq.com` |
| `DAILY_ATTACHMENT_ARCHIVE_DIR` | 可选本地每日归档目录 |
| `NUTSTORE_WEBDAV_URL`、`NUTSTORE_WEBDAV_USER`、`NUTSTORE_WEBDAV_APP_PASSWORD` | WebDAV 连接 |
| `NUTSTORE_REMOTE_DAILY_ARCHIVE_DIR`、`NUTSTORE_REMOTE_MONTHLY_ARCHIVE_DIR`、`NUTSTORE_REMOTE_AUTO_LIST_FILE` | 远程归档和当前自动清单 |

月度工具的本地目录等参数见 [月度契约](docs/MONTHLY_SUMMARY_SPEC.md)；
不凭文档旧收件人数、目录或历史 secret 快照推断现值。

## 日报与月度验收

[每日工作流](.github/workflows/run-daily.yml) 没有 cron，仅支持 `workflow_dispatch`。
IMAP 不可达时会跳过流程；仅 `020 Email download.py` 失败时可软失败放行。
因此 Actions success 不足以证明日报完成，至少检查：

- artifact `pipeline-run-report` 中 `run-report.json` 的 `success: true`、
  未被 `skipped` 跳过、`failed_steps: []`。
- 发信日志出现“邮件发送成功”，并核对最终内容与目标收件人；用户确认收到强于发送日志。
- 如验收归档，还需独立核对目标目录及文件，不能由发信成功推断归档成功。

[月初工作流](.github/workflows/monthly-closeout.yml) 核验上一月月末附件与每日归档的语义指纹。
只有一致才写自动月末归档和报告；不一致只发人工核验通知，不覆盖月末文件。
缺少归档输入属于“尚未核验”，不是“指纹不一致”。最近已知缺口和验证方式见
[当前交接](docs/CURRENT_HANDOFF.md)。

## 维护

> 邮件图片样式优化：`050 image.py` 使用 LibreOffice Calc 按 Excel 原格式渲染 `库存表`，裁白边并将最终 PNG 宽度限制为 1800px；运行依赖和标题口径见 `docs/PIPELINE_FLOW.md`、`docs/CURRENT_HANDOFF.md`。

编排改 `main.py` / `pipeline/`，业务步骤改 `script/`，独立工具改 `tools/`。
步骤和输入输出变化时同步 STEP CARD、现有契约和流程图；只更新受影响文档，
不为无影响的改动增加记录。提交、推送、PR、真实运行与远程配置修改需要明确授权。

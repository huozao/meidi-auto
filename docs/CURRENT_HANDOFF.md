# 美的库存自动化：当前交接状态

> 更新日期：2026-09-13。新会话处理本项目时先读本文，再按需阅读
> `README.md`、`docs/PIPELINE_FLOW.md` 和 `docs/MONTHLY_SUMMARY_SPEC.md`。

## 代码、数据与执行边界

- **唯一代码源**：GitHub 仓库 `huozao/meidi-auto` 的 `main`。WSL 工作目录只作为开发、排查和本地验证副本；不得以本地未推送代码作为生产版本。
- **当前上线基线**：`94cef848b683b3acbd59e41882674c68adf2f0fa`（PR #11，2026-09-13）。该提交已包含 Excel 合计、公司领用集中度、日报 HTML 展示及自动清单口径。
- **业务数据不入 Git**：邮件附件、月末快照、Excel/HTML 报告和坚果云同步文件均是运行数据；邮箱与 WebDAV 凭据只存在 WSL `.env` 或 GitHub Actions Secrets。
- **数据留存**：每日附件归档到 `NUTSTORE_REMOTE_DAILY_ARCHIVE_DIR/YYYYMM/`；月度快照、报告、当前自动清单和按月副本归入 `NUTSTORE_REMOTE_MONTHLY_ARCHIVE_DIR`。Windows 坚果云同步目录与 WebDAV 是同一业务留存体系，但不是代码源。

## 当前触发与验收

### 每日日报

- 工作流：`Run MeidiAuto Pipeline`（`.github/workflows/run-daily.yml`）。仅保留 `workflow_dispatch`，没有 GitHub cron。
- 正常自动链路：Gmail Watch -> Pub/Sub -> `gmail-watcher` Cloud Run -> GitHub `workflow_dispatch` -> 本仓工作流；人工复核也可以在 Actions 页面或用 `gh workflow run` 手动发起。
- 最近实测：运行 `34734283700`，使用上线 SHA `94cef848`，完成时间约 83 秒；从坚果云读取自动清单工作表 `2608`，生成 HTML，日志出现“邮件发送成功”，artifact `run-report.json` 为 `success: true` 且 `failed_steps: []`。
- **验收不能只看 Actions 绿灯**：该工作流对仅 `020 Email download.py` 失败存在软失败放行。必须同时检查 artifact 的 `run-report.json`、`failed_steps` 以及日志中的“邮件发送成功”。

推荐复核命令（不包含任何凭据）：

```bash
gh run view <run-id> --repo huozao/meidi-auto --log
gh run download <run-id> --repo huozao/meidi-auto --name pipeline-run-report
```

### 月初月末核验

- 工作流：`Meidi monthly closeout`（`.github/workflows/monthly-closeout.yml`）。北京时间每月 1 日 00:02 运行；其他日期的 cron 会被 gate 跳过，亦可 `workflow_dispatch` 手动核验。
- 判据：取上月最后一天 24:00 前的最新目标邮件附件；其与每日归档语义指纹一致才会写入自动月末归档、月度报告和新的自动清单。不一致时仅发人工核验通知，不覆盖已有文件。
- **当前待跟进项**：手动回溯运行 `34731737697`（目标 `2026-08`）在比较前失败，原因是 WebDAV 的每日归档目录 `NUTSTORE_REMOTE_DAILY_ARCHIVE_DIR/202608` 返回 404。此结果表示缺少核验输入，**不是**“月末文件不一致”。在下一次真实月初核验前，应确认上一月目录已存在，且含有月末当天的 Excel；未完成一次真实月份的“指纹一致 + 报告回写”前，不应宣称月初自动闭环已通过生产验证。

## 已确认的业务口径

- 月末快照合格的主判据是工作簿 `库存表` 物料表头上方有目标月最后一天日期（例如 `2026年08月31日`）。当天没有出入库业务不是失败，只记录为数据质量提示。
- 自动清单由最近三个**完整**月份生成：外应存和家应存各为近三个月出库总量 `/ 12`（三月周均）；月计划为 `/ 3`（三月月均）。人工维护的“备注”从上一版清单继承，其余数值每月重算。
- 每日日报只读当前自动清单；自动清单不存在或读取失败时禁止回退旧 `script/data/list.xlsx`，邮件必须明确提示“无法计算”。
- 公司识别优先用 `出入库明细表` 的备注；`MA1111` 归“重庆工厂”，`MA1141`、`MS1121` 分别按各自公司名统计。其他未识别记录保留为“未备注”，不得静默排除。
- 每日 HTML 的“物料出库趋势”按近三月总出库量倒序；“前五大领用公司及物料明细”展示公司及其物料月度用量和今日变化。两表底部直接追加合计行；公司表再追加全部领用出库合计和前五大公司领用集中度（分母含“未备注”领用）。
- 月度 Excel 的可加数量表均在末尾追加合计；公司月度/年度汇总含“全部领用出库合计”和占比，避免从 HTML 重算口径。

## 新会话的最小检查清单

1. `git status --short`、`git log -1` 和 `git ls-remote origin refs/heads/main`：本地 `main` 必须与 GitHub 远端 SHA 一致。
2. 修改代码前运行 `.venv/bin/python -m unittest discover -s tests -v`；提交前再执行 `py_compile` 和 `git diff --check`。
3. 若验证日报，按上文的业务验收三要素检查，不以 Actions 结论代替邮件业务结果。
4. 若验证月初任务，先确认上一月 WebDAV 每日归档目录及月末 Excel 已存在，再触发；缺目录应先报告输入缺失，不要重试或覆盖月末快照。
5. 所有代码与文档改动经测试后提交、推送并合并回 GitHub `main`；不提交 `.env`、附件、WebDAV 下载、报告或浏览器/运行缓存。

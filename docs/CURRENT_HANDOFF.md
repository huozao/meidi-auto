# 美的库存自动化：当前交接状态

> 更新日期：2026-09-19。新会话处理本项目时先读本文，再按需阅读
> `README.md`、`docs/PIPELINE_FLOW.md` 和 `docs/MONTHLY_SUMMARY_SPEC.md`。

本文的 SHA、运行记录与资源缺口是上述日期的交接快照，后续线上状态需重新取证。
本文不授予提交、推送、合并或真实发信权限；纯本地任务不要求访问 GitHub、邮箱或 WebDAV。

## 代码、数据与执行边界

- **唯一代码源**：GitHub 仓库 `huozao/meidi-auto` 的 `main`。WSL 工作目录只作为开发、排查和本地验证副本；不得以本地未推送代码作为生产版本。
- **当前上线基线**：`f5a1379808a6683b04ea2ab859296ac5a7a00ef6`（GitHub `main`，2026-09-19）。该提交包含 PR #14 的库存两周均值与邮件收件时间改动；GitHub `main` 是生产唯一代码源。
- **业务数据不入 Git**：邮件附件、月末快照、Excel/HTML 报告和坚果云同步文件均是运行数据；邮箱与 WebDAV 凭据只存在 WSL `.env` 或 GitHub Actions Secrets。
- **数据留存**：每日附件归档到 `NUTSTORE_REMOTE_DAILY_ARCHIVE_DIR/YYYYMM/`；月度快照、报告、当前自动清单和按月副本归入 `NUTSTORE_REMOTE_MONTHLY_ARCHIVE_DIR`。Windows 坚果云同步目录与 WebDAV 是同一业务留存体系，但不是代码源。

## 当前触发与验收

### 每日日报

- 工作流：`Run MeidiAuto Pipeline`（`.github/workflows/run-daily.yml`）。仅保留 `workflow_dispatch`，没有 GitHub cron。
- 正常自动链路：Gmail Watch -> Pub/Sub -> `gmail-watcher` Cloud Run -> GitHub `workflow_dispatch` -> 本仓工作流；人工复核也可以在 Actions 页面或用 `gh workflow run` 手动发起。
- 最近实测：运行 `34844719311`（2026-09-14，使用 SHA `69ea637`），耗时 76.79 秒；读取当日文件 `总库存20260914_204108.xlsx`，生成 HTML，日志出现“邮件发送成功”（配置收件人数 3），并完成坚果云 3 个文件归档。artifact `run-report.json` 为 `success: true` 且 `failed_steps: []`。该记录证明发送端成功，不等同于收件人已读或已收到。
- **验收不能只看 Actions 绿灯**：该工作流对仅 `020 Email download.py` 失败存在软失败放行。必须同时检查 artifact 的 `run-report.json`、`failed_steps` 以及日志中的“邮件发送成功”。

### 2026-09-19 库存口径与邮件收件时间交付

- PR [#14](https://github.com/huozao/meidi-auto/pull/14) 已合并，合并提交为 `f5a1379808a6683b04ea2ab859296ac5a7a00ef6`。外应存、家应存分别按近三个月出库总量 `/ 6` 计算两周均值，两处合计 `/ 3` 覆盖一个月均用量；日报 HTML 优先读取 `mail_meta.json` 中两封来源邮件的收件时间并展示到秒，缺失时回退 H3/M3。
- 首次生产运行 [35421736354](https://github.com/huozao/meidi-auto/actions/runs/35421736354) 于 2026-09-19（UTC）完成，耗时 57.25 秒；artifact `run-report.json` 为 `success: true`、`failed_steps: []`、`skipped` 未启用，11 个步骤均执行成功。
- 运行日志确认：`RECIPIENT_EMAILS` 配置为 3 个收件人，出现“邮件发送成功”；坚果云归档完成 3 个文件。发送日志/归档日志证明发送端和归档端成功，不等同于收件人已读或实际收到；用户计划当晚观察邮件展示。
- 本次生产运行使用主题“美的库存及出入库日报｜2026-09-18｜库存、出入库及月计划”；邮件正文和附件为运行时产物，不入 Git。后续如需核对 HTML 的秒级时间，应以收件人实际邮件正文为准。

推荐复核命令（不包含任何凭据）：

```bash
gh run view <run-id> --repo huozao/meidi-auto --log
gh run download <run-id> --repo huozao/meidi-auto --name pipeline-run-report
```

### 月初月末核验

- 工作流：`Meidi monthly closeout`（`.github/workflows/monthly-closeout.yml`）。北京时间每月 1 日 00:02 运行；其他日期的 cron 会被 gate 跳过，亦可 `workflow_dispatch` 手动核验。
- 判据：取上月最后一天 24:00 前的最新目标邮件附件；其与每日归档语义指纹一致才会写入自动月末归档、月度报告和新的自动清单。不一致时仅发人工核验通知，不覆盖已有文件。
- **当前待跟进项**：手动回溯运行 `34731737697`（目标 `2026-08`）在比较前失败，原因是 WebDAV 的每日归档目录 `NUTSTORE_REMOTE_DAILY_ARCHIVE_DIR/202608` 返回 404。此结果表示缺少核验输入，**不是**“月末文件不一致”。在下一次真实月初核验前，应确认上一月目录已存在，且含有月末当天的 Excel；未完成一次真实月份的“指纹一致 + 报告回写”前，不应宣称月初自动闭环已通过生产验证。

### 2026-09-16 邮件标题与图片附件交接

- 用户已确认标题：`美的库存及出入库日报｜YYYY-MM-DD｜库存、出入库及月计划`；日期取 `库存表!H3`，不再从 PNG 文件名拼接。
- 用户已确认图片方案：`050 image.py` 使用 LibreOffice Calc 按 `库存表` 的 Excel 样式导出，渲染副本只保留 `A1:T末行`，裁白边并限制最终 PNG 宽度为 `1800px`；原始 Excel 不被改写。
- GitHub runner 需要 `libreoffice-calc`、`fonts-noto-cjk`、`poppler-utils`，Python 依赖使用 `Pillow`；安装步骤在 `.github/workflows/run-daily.yml`。
- 本地案例使用 2026-09-15 最新邮件附件验证：标题为 `美的库存及出入库日报｜2026-09-15｜库存、出入库及月计划`，图片为 `1800×985px`，中文、合并表头、颜色、数字格式和合计行均可读。该案例文件在 Windows 业务数据目录，不入 Git。
- 本次代码交付为 PR [#13](https://github.com/huozao/meidi-auto/pull/13)，核心代码 commit `f97b56015ff27d31ce9e804bae2de89fe6fa0f09`，交接文档随 PR 一并更新；合并后首次生产日报需重新核对邮件主题、图片附件和收件人实际收到的内容。
- 首轮生产观察至少检查：Actions 日志的 `邮件发送成功`、artifact `run-report.json` 的 `success` / `failed_steps`、实际邮件主题和图片清晰度；Actions 绿灯不等于业务验收完成。若失败，先保留运行产物和日志，再通过 PR 回滚，不直接修改生产数据。

## 已确认的业务口径

- 月末快照合格的主判据是工作簿 `库存表` 物料表头上方有目标月最后一天日期（例如 `2026年08月31日`）。当天没有出入库业务不是失败，只记录为数据质量提示。
- 自动清单由最近三个**完整**月份生成：外应存和家应存各为近三个月出库总量 `/ 6`（三月两周均），两处合计覆盖一个月平均用量；月计划为 `/ 3`（三月月均）。人工维护的“备注”从上一版清单继承，其余数值每月重算。
- 每日日报只读当前自动清单；自动清单不存在或读取失败时禁止回退旧 `script/data/list.xlsx`，邮件必须明确提示“无法计算”。
- 公司识别优先用 `出入库明细表` 的备注；`MA1111` 归“重庆工厂”，`MA1141`、`MS1121` 分别按各自公司名统计。其他未识别记录保留为“未备注”，不得静默排除。
- 每日 HTML 的“物料出库趋势”按近三月总出库量倒序；“前五大领用公司及物料明细”展示公司及其物料月度用量和今日变化。两表底部直接追加合计行；公司表再追加全部领用出库合计和前五大公司领用集中度（分母含“未备注”领用）。
- 每日 HTML 展示方案（`script/050 mailtxt.py`、`pipeline/usage_trends.py`）：不再输出颜色提示、统计范围、取样逻辑等解释文字；优先使用 `mail_meta.json` 的两封来源邮件收件时间，展示到 `YYYY-MM-DD HH:MM:SS`，缺少元数据时回退 H3/M3；汇总左组顺序为出库、入库、外仓库存总量、家里库存总量，右组顺序为月计划、月计划预估还有要发货、月计划缺口排产。物料出库趋势和前五大公司表的月份按最近到最远排列，最近月份单元格仅在今日有非零变化时追加“今日 ±数量”，为零时不显示提示；月度契约中的同名说明已同步更新。
- 月度 Excel 的可加数量表均在末尾追加合计；公司月度/年度汇总含“全部领用出库合计”和占比，避免从 HTML 重算口径。

## 新会话的最小检查清单

1. 本地任务先看 `git status --short`、`git log -1`，保留其他会话改动。准备已授权的交付或线上验证时，再用 `git ls-remote origin refs/heads/main` 核对远端基线；开发分支或未推送改动不能当作线上版本。
2. 修改代码前运行 `.venv/bin/python -m unittest discover -s tests -v`；提交前再执行 `py_compile` 和 `git diff --check`。
3. 若验证日报，按上文的业务验收三要素检查，不以 Actions 结论代替邮件业务结果。
4. 若验证月初任务，先确认上一月 WebDAV 每日归档目录及月末 Excel 已存在，再触发；缺目录应先报告输入缺失，不要重试或覆盖月末快照。
5. 只有获得明确交付授权后，才将代码与受影响文档经测试、提交、推送并通过 PR 合并回 GitHub `main`；本地修改任务不自动进入交付流程。不提交 `.env`、附件、WebDAV 下载、报告或浏览器/运行缓存。

## 当前工作区边界（2026-09-19）

- GitHub `main` 当前远端 SHA 为 `f5a1379808a6683b04ea2ab859296ac5a7a00ef6`，PR #14 已合并；生产运行 `35421736354` 已完成。后续 AI 仍应先执行 `git ls-remote origin refs/heads/main` 和 `git status --short`，不要把本地副本当作生产代码源。
- 原始工作区仍保留一组未提交的既有指导/历史文档改动：`AGENTS.md`、`README.md`、`CLAUDE.md`、`docs/development.md`、`docs/migrations/`，以及本次早期开发副本的业务文件改动。它们未进入 PR #14，也未进入 GitHub `main`；不得 reset、clean 或未经核对混入后续提交。若后续需要交付这些文档，必须单独审阅并走 PR。
- 生产 artifact 仅保留 `pipeline-run-report`；Excel、HTML、邮件附件和 WebDAV 文件是运行时业务数据，不入 Git。发送日志不能替代收件人实际收到的证明。

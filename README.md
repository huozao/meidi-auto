# meidi-auto

用于在 GitHub Actions 或本地执行的库存自动化流水线。

## 目录结构

- `main.py`：统一编排入口，负责参数解析与调度。
- `pipeline/`：编排层模块（步骤定义、数据模型、产物校验）。
  - `pipeline/models.py`：步骤数据模型。
  - `pipeline/steps.py`：生产步骤与清理步骤清单。
  - `pipeline/validators.py`：步骤脚本存在性与关键产物校验。
  - `pipeline/io_utils.py`：脚本共享的路径与文件发现工具。
  - `pipeline/flow_map.py`：业务阶段与常见需求修改导航。
- `script/`：业务脚本（下载邮件、合并 Excel、计算与着色、生成 HTML、发送邮件）。
- `.github/workflows/run-daily.yml`：标准化 CI 运行工作流（由 `gmail-watcher` 触发，也支持手动运行）。
- `requirements.txt`：Python 依赖。
- `docs/PIPELINE_FLOW.md`：主流程一页图（Mermaid）+ 阅读顺序。
- `docs/AI_FRIENDLY_REDESIGN.md`：AI 友好型重构说明（主线、边界、改动导航）。

## 运行方式

### 本地运行

```bash
python -m pip install -r requirements.txt
python main.py --check
python main.py --data-dir data --stop-on-error --report-file data/run-report.json
```

推荐按“由浅到深”验证：

1. `python main.py --check`：检查脚本是否齐全、环境变量是否就绪。
2. `python main.py --dry-run`：确认执行顺序和参数无误。
3. `python main.py --data-dir data-local --stop-on-error --report-file data-local/run-report.json`：做一次完整联调并保留报告。

若要按阶段逐步核对结果（推荐）：

```bash
python main.py --list-steps
python main.py --only-step 020 --data-dir data-local
python main.py --only-step 021 --data-dir data-local
python main.py --only-step 030 --data-dir data-local
```

说明：
- `--only-step` 可用文件名前缀/关键字（如 `030`、`041 operation.py`）；
- 适合你逐步检查每个阶段输出是否符合预期；
- 不建议通过注释代码来“跳过步骤”，容易引入误差和忘记恢复。

## 新手说明：`--dry-run` 到底是什么？

你可以把 `--dry-run` 理解成“**演习模式**”：

- 会把“将要执行哪些步骤”打印出来；
- **不会真正执行**下载邮件、处理 Excel、发邮件等动作；
- 适合第一次上手时确认流程是否正确。

对比：

- `python main.py --dry-run`：只看计划，不改任何文件；
- `python main.py --check`：检查环境变量和脚本是否存在；
- `python main.py ...`（不带 `--dry-run`）：真正执行流程。

建议新手顺序：先 `--check`，再 `--dry-run`，最后正式执行。

如果你要手动清理测试产物或自动运行后的冗余文件，可用：

- `python main.py --clean-only`：只执行 `script/010 clean.py`；
- `python main.py --clean-after-run --stop-on-error`：主流程结束后追加执行 `010 clean.py`。

目录说明（你截图那种结构是正常的）：

- `script/`：脚本目录；
- `data/`：主流程输出目录（默认）；
- `script/data/`：部分历史脚本可能使用的目录。

当前 `010 clean.py` 会优先清理 `main.py` 传入的数据目录，并兼容清理 `data/` 与 `script/data/`。

运行结果重点看 3 个点：

- 进程退出码是否为 `0`；
- 汇总区是否出现失败步骤；
- `run-report.json` 中 `success` 是否为 `true`、`failed_steps` 是否为空数组。

另外，`main.py` 对关键步骤增加了基础产物校验（如 020 必须生成 `mail_meta.json` 和 `存量查询*.xlsx`、021 必须生成 `总库存*.xlsx`），可避免子脚本误返回成功时导致后续级联失败。

### Windows 编码问题说明（UnicodeDecodeError: gbk）

若在 Windows/PyCharm 中看到 `UnicodeDecodeError: 'gbk' codec can't decode ...`，通常是子脚本输出编码与父进程解码编码不一致导致。当前 `main.py` 已改为二进制捕获并自动尝试 `utf-8`/`gbk` 解码，减少该类报错对联调的干扰。

### Docker 运行（推荐用于环境一致性验证）

可以。把项目拉到 Docker 后同样能跑，且更接近 CI 环境。示例：

```bash
docker run --rm -it \
  -v "$PWD":/app \
  -w /app \
  -e EMAIL_ADDRESS_QQ=xxx \
  -e EMAIL_PASSWORD_QQ=xxx \
  -e RECIPIENT_EMAILS="a@example.com,b@example.com" \
  -e IMAP_SERVER=imap.qq.com \
  python:3.12 bash -lc "pip install -r requirements.txt && python main.py --check && python main.py --data-dir data-docker --stop-on-error --report-file data-docker/run-report.json"
```

### 新手版：Docker 从拉取到运行（一步一步）

1. **拉取 Python 镜像**

```bash
docker pull python:3.12
```

2. **进入项目目录**（里面要有 `main.py` 和 `requirements.txt`）

```bash
cd meidi-auto
```

3. **运行容器并执行流水线**

```bash
docker run --rm -it \
  -v "$PWD":/app \
  -w /app \
  -e EMAIL_ADDRESS_QQ=你的邮箱 \
  -e EMAIL_PASSWORD_QQ=你的授权码 \
  -e RECIPIENT_EMAILS="收件人1,收件人2" \
  -e IMAP_SERVER=imap.qq.com \
  python:3.12 bash -lc "pip install -r requirements.txt && python main.py --check && python main.py --dry-run"
```

4. **确认无误后再执行正式跑**

```bash
docker run --rm -it \
  -v "$PWD":/app \
  -w /app \
  -e EMAIL_ADDRESS_QQ=你的邮箱 \
  -e EMAIL_PASSWORD_QQ=你的授权码 \
  -e RECIPIENT_EMAILS="收件人1,收件人2" \
  -e IMAP_SERVER=imap.qq.com \
  python:3.12 bash -lc "pip install -r requirements.txt && python main.py --data-dir data-docker --stop-on-error --clean-after-run --report-file data-docker/run-report.json"
```

5. **查看结果文件**

- `data-docker/run-report.json`：最终成功/失败报告；
- 若启用了 `--clean-after-run`，会在主流程后执行 `010 clean.py` 做清理。

若需要在容器中一并清理产物，可在末尾追加：

```bash
python main.py --data-dir data-docker --stop-on-error --clean-after-run --report-file data-docker/run-report.json
```

### GitHub Actions 自动运行

仓库已提供 `.github/workflows/run-daily.yml`，支持以下触发方式：

- `workflow_dispatch`：由 `gmail-watcher` 在收到匹配邮件后触发，也可以在 Actions 页面手动点击运行。

#### 当前触发边界（2026-09-11）

- PR #5（合并提交 `e23205c`）已移除 `schedule`；此工作流不再按 GitHub cron 自动运行。
- `workflow_dispatch` 保留不变，`gmail-watcher` 仍通过匹配关键词后的 GitHub dispatch 触发本流水线。
- 2026-09-11 11:44（北京时间）收到的邮件来自旧版本的 `schedule` 运行 `34559512117`；该运行已完成邮件发送。不要把这次历史运行误判为当前仍有定时配置。
- 排查时先看 Actions 的 `event`（`schedule` 或 `workflow_dispatch`），再看 artifact 中 `run-report.json` 的 `failed_steps` 和日志里的 `✅ 邮件发送成功`。

首次启用前，请在仓库 `Settings -> Secrets and variables -> Actions` 中配置：

- `EMAIL_ADDRESS_QQ`
- `EMAIL_PASSWORD_QQ`（或兼容旧变量 `EMAIL_PASSWOR_QQ`）
- `RECIPIENT_EMAILS`（必需；逗号分隔多个收件人）
- `IMAP_SERVER`（可选）

> 若 `IMAP_SERVER` Secret 留空，workflow 会自动回退到 `imap.qq.com`。

### GitHub Actions 跑完后，`data/` 会不会留痕？

默认不会长期留痕在仓库里：

- GitHub Actions Runner 是临时环境，任务结束后工作目录会销毁；
- `data/` 里的文件只在本次任务内存在；
- 只有显式 `upload-artifact` 的内容会被保留（当前主要保留运行报告）。

所以把运行中间产物放在 `data/` 是合理的；需要长期保存时再按需上传 artifact 即可。

### GitHub Actions 常见失败：`Connection refused`（IMAP 连接被拒）

若日志出现：

- `获取邮件失败: [Errno 111] Connection refused`

通常不是代码逻辑错误，而是邮箱服务侧对 GitHub 公网 Runner 网络做了限制（或端口策略、白名单策略）。

建议按优先级处理：

1. 先确认 `IMAP_SERVER`、邮箱授权码、IMAP 开关都正确；
2. 在本地网络验证同一账号可连通；
3. 若仅 GitHub Hosted Runner 失败，改用 **self-hosted runner**（公司/本地固定出口 IP）执行；
4. 或在邮箱侧配置允许策略（如有白名单能力）。

说明：当前流水线已配置 `--stop-on-error`，020 步骤连不上邮箱会直接失败并停止，避免后续步骤在无输入文件时继续报级联错误。
另外，`run-daily.yml` 已增加 IMAP 连通性探测：若 GitHub Hosted Runner 无法连通 IMAP，会跳过生产步骤并输出一份“skipped”报告，避免误报为代码逻辑错误。
对于偶发的 Hosted Runner 网络抖动，workflow 还会在评估阶段把“仅 020 失败”的场景按软失败处理（保留报告但不直接判整条流水线红灯）。

### 为什么“我贴的程序”看起来没问题，但 GitHub 还是不稳定？

因为你贴的是**代码逻辑**，而 GitHub 失败多数是**运行环境差异**：

- 本地网络可访问 IMAP，不代表 GitHub Hosted Runner 也能访问；
- 邮箱服务可能对公网机房 IP、端口、TLS 策略有限制；
- 同样代码在“可连通环境”会成功，在“不可连通环境”会失败。

所以这不是“代码一定错了”，而是“代码 + 运行环境”共同决定结果。

为便于排查配置问题，020/051 步骤会打印脱敏后的环境变量摘要（如邮箱打码、密码长度、收件人数），不会输出明文授权码。

### 本次问题复盘（为什么这个版本恢复正常）

这次能正常登录 QQ 邮箱并发送邮件，主要是因为以下几个问题被同时修复：

1. **`IMAP_SERVER` 空值问题**  
   之前 GitHub Secrets 中该值为空时，会把空字符串注入流程，导致探测报 `name or service not known`。现在 workflow 已改为：空值自动回退到 `imap.qq.com`。

2. **收件人配置方式统一**  
   收件人改为从 `RECIPIENT_EMAILS` 读取，避免脚本内硬编码与环境配置不一致。

3. **邮箱步骤失败语义更明确**  
   020 在未获取到 HTML 时会明确失败，不再“假成功”进入后续步骤，排错路径更清晰。

4. **日志可观测性提升**  
   现在会打印脱敏后的关键环境摘要（账号打码、密码长度、收件人数），便于快速判断是不是 Secrets 注入错误。

### 本地调试建议：断点 / PyCharm / Docker 怎么选

推荐优先级：

1. **先用 `--only-step` 分阶段运行**（最快）；
2. **再在 PyCharm 对单个脚本打断点**（最直观）；
3. **Docker 主要用于环境一致性验证**，断点调试成本更高（需远程调试配置）。

结论：

- PyCharm：非常适合加断点调试；
- Docker：可以调试但配置复杂，先不作为首选；
- 注释 `main.py` 某些行：不推荐，容易改乱流程，建议用参数控制（`--only-step` / `--clean-only`）。

### GitHub Actions 与本地是否可同时运行

可以同时进行：

- 远端 GitHub Actions 在 GitHub Runner 上执行；
- 本地运行在你自己的电脑上执行。

两者互不阻塞，但建议注意以下事项：

- 两边都在读同一邮箱时，可能出现重复下载/重复处理；
- 建议本地测试时使用独立 `--data-dir`（例如 `data-local`）；
- 若需避免同一时段重复发送邮件，可将本地运行设置为 `--dry-run` 先验证流程。

### 仅查看执行计划

```bash
python main.py --dry-run
```

## 必需环境变量

- `EMAIL_ADDRESS_QQ`
- `EMAIL_PASSWORD_QQ`（兼容历史变量 `EMAIL_PASSWOR_QQ`）
- `RECIPIENT_EMAILS`（必需，例：`a@example.com,b@example.com`）
- `IMAP_SERVER`（可选，默认 `imap.qq.com`）

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
python main.py --dry-run
```

4. 再执行完整链路（建议先用独立数据目录）：

```bash
python main.py --data-dir data-local --stop-on-error --report-file data-local/run-report.json
```

5. 若要云端验证，把同名变量配置到 GitHub Secrets，然后手动触发 `run-daily.yml` 的 `workflow_dispatch`。

可按下列映射填写：

- `.env` 的 `EMAIL_ADDRESS_QQ=xxx` -> GitHub Secret 名称：`EMAIL_ADDRESS_QQ`
- `.env` 的 `EMAIL_PASSWORD_QQ=xxx` -> GitHub Secret 名称：`EMAIL_PASSWORD_QQ`
- `.env` 的 `RECIPIENT_EMAILS=a@example.com,b@example.com` -> GitHub Secret 名称：`RECIPIENT_EMAILS`
- `.env` 的 `IMAP_SERVER=imap.qq.com` -> GitHub Secret 名称：`IMAP_SERVER`（可选）

## 设计原则

- 以 `main.py` 作为单一入口，避免多入口导致流程分叉。
- 工作流统一使用一个 YAML，避免重复/失效配置。
- 数据目录集中为 `data/`，所有脚本通过参数共享同一输出目录。

## 代码质量简评

整体评价：**中上（7.5/10）**，已经具备可维护流水线项目的核心骨架。

优点：

- 有统一编排入口（`main.py`），步骤顺序清晰，避免“脚本各跑各的”。  
- 参数化较完整（`--check` / `--dry-run` / `--stop-on-error` / `--report-file`），便于本地与 CI 共用。  
- 有最基本的可观测性：步骤耗时、失败列表、JSON 报告。  
- 兼容历史环境变量名（`EMAIL_PASSWORD_QQ|EMAIL_PASSWOR_QQ`），降低迁移成本。  

可改进点：

- 当前步骤间通过“文件约定”耦合，建议逐步补充输入/输出契约文档。  
- 缺少自动化测试（至少应补 1 组 smoke test 和 1 组失败场景测试）。  
- 失败重试策略还比较薄弱（网络/邮箱抖动时可加可配置重试）。  
- 日志结构化程度一般，后续可考虑标准 logging + 统一字段。  

## `main.py` 为什么要这样写

`main.py` 采用“编排器（orchestrator）”模式，核心目的不是承载业务细节，而是把多脚本流水线的**执行顺序、错误策略、运行入口**统一起来：

1. **单一入口，减少分叉**  
   通过 `PIPELINE_STEPS` 固化步骤顺序，避免本地、CI、临时脚本三套流程不一致。

2. **检查与执行分离**  
   `--check` 只做前置校验，`--dry-run` 只看计划，正式执行再真正跑子脚本，便于快速定位问题来源。

3. **统一错误处理策略**  
   每步都走同一套 `run_step`，可在 `--stop-on-error` 下快速失败，也可完整跑完后统一汇总失败列表。

4. **统一数据目录与报告出口**  
   所有步骤共享 `--data-dir`，并可输出 `--report-file`，方便本地复现与 CI 归档。

5. **兼容历史系统**  
   环境变量读取支持新旧键名，既保证演进，也避免“一次性切换”造成线上失败。


## 兼容说明

- 默认生产链路以 `main.py` + `run-daily.yml` 为准。
- 已移除失效的历史 workflow，避免在 GitHub Actions 中触发不存在脚本导致失败。
- 已删除冗余邮件脚本 `script/052 send email.py`，统一由 `script/051 Send an email.py` 负责发信。


## 运行前检查

```bash
python main.py --check
```

用于验证：
- 核心脚本是否存在；
- 必需环境变量是否已配置。

## 运行报告

可通过 `--report-file` 输出一次运行的 JSON 报告，便于追踪失败步骤。

## 架构评估与演进建议（讨论版）

当前“多子脚本串行处理 Excel”的方案**可用且直观**，适合业务规则经常变化、由非开发同事共同维护的阶段。

从运行日志看，现状优势：

- 每一步输入/输出都可见，定位问题快；
- 单步失败可快速重跑，不必全链路重来；
- Excel 产物天然可人工复核，利于业务确认。

但中长期风险也明显：

- 脚本之间靠“文件名约定 + 覆盖写回”传递状态，耦合较强；
- 同一文件多次读写（021/030/032/033/041/042）对性能和一致性不友好；
- 业务规则散落在多个脚本，回归测试成本高。

建议按“低风险渐进”改造，而不是一次重写：

1. 先统一数据模型：把每步关键字段落成 `data/intermediate/*.parquet|csv`；
2. 再把 030/032/033/041 的纯计算逻辑收敛到一个 `transform` 模块；
3. 最后保留 042（着色）与 051（发邮件）作为独立输出层。

这样能保留你现在的可读性与可回滚能力，同时明显降低“脚本链条越拉越长”的维护风险。

## 现在这版是否更便于维护/排错/扩展列？

结论：**是的，比之前明显更好**，原因如下：

1. **维护性更好**  
   编排层已拆到 `pipeline/`（步骤定义、校验、IO 工具），不再全部堆在 `main.py`，改动边界更清晰。

2. **排错更快**  
   增加了步骤级产物校验与分步骤执行参数（`--only-step`），可以快速定位是“哪一步没产物”还是“哪一步计算异常”。

3. **扩展新内容列更可控**  
   你可以只改目标业务脚本（例如 033/041 的列映射配置），再用 `--only-step` 做小范围回归，不必每次全链路回归。

建议新增列时按这个顺序：

1. 在对应脚本的配置区新增列映射（优先配置化，不写死逻辑）；
2. 增加/更新该脚本的日志打印（至少打印列名、写入行数）；
3. 本地用 `python main.py --only-step <步骤号> --data-dir data-local` 验证；
4. 再跑 `--dry-run` + 全链路；
5. 最后确认 `run-report.json` 与最终 Excel 结果。

## 常见问题：出现 `Accept current/incoming/both changes` 是什么？

这是 **Git 合并冲突（merge conflict）**，不是运行错误。通常出现在你本地分支和远端分支都改了同一段代码时。

- `Accept current change`：保留你当前分支的代码；
- `Accept incoming change`：保留对方（要合并进来）分支的代码；
- `Accept both changes`：两边都保留（随后通常还要手动整理一次）。

你**不需要去网页上“确认提交代码”**，而是要先在本地把冲突处理完，再提交：

```bash
git add <冲突文件>
git commit -m "resolve merge conflict"
git push
```

建议流程（新手版）：

1. 先看冲突文件里 `<<<<<<<`, `=======`, `>>>>>>>` 三段标记；
2. 选 current / incoming / both 其中一种；
3. 删除冲突标记并保证代码可运行；
4. 再执行一次 `python main.py --dry-run` 和 `python main.py --check`；
5. 最后 `git add` + `git commit` + `git push`。

### 冲突时到底该怎么选（直接结论）

按这个顺序判断即可：

1. **如果 incoming 里有你刚新增的参数/功能（例如 `--clean-only`、`--clean-after-run`）而 current 没有：选 `Accept incoming change`。**
2. **如果 current 和 incoming 各自改了不同内容：选 `Accept both changes`，然后手动删重复行，保留两边有效改动。**
3. **如果 incoming 明显是旧版本（把中文日志、新参数、编码修复都删掉）：选 `Accept current change`。**

针对你截图那段 `parse_args()`，通常应优先保留包含 `--clean-only` 与 `--clean-after-run` 的版本；如果另一边还有你需要的参数，就用 `Accept both` 后手工合并成一份完整参数列表。

## 项目评估与维护地图（新增）

为了方便后续维护者快速定位“该改哪个模块”，仓库新增了维护文档：

- `docs/MAINTENANCE_MAP.md`

文档包含：

1. 主流程每个程序模块的功能、输入输出与依赖；
2. 疑似冗余/历史脚本清单（以及为什么建议归档）；
3. 日常维护时的“改哪里”速查与最小回归验证路径。

建议修改代码前先读该文档再动手。

### 维护表格（模块清单）自动生成

如果你想长期维护“模块功能表”并避免手工维护出错，使用下面命令：

```bash
python tools/generate_module_catalog.py
python tools/generate_module_catalog.py --check
```

说明：
- 源数据：`docs/module_catalog.json`
- 自动产物：
  - `docs/MODULE_CATALOG.md`（给人看）
  - `docs/module_catalog.csv`（给表格工具/二次分析用）
- `--check` 适合在 CI 中做“文档是否过期”校验。

### 结构评分与优化建议（2026-04-01）

当前“主程序串联子程序”结构评分：**8.0 / 10**（已完成中期优化第一阶段）。

优点：
- `main.py` 统一调度、支持 `--check/--dry-run/--only-step`；
- 步骤顺序集中在 `pipeline/steps.py`，可维护性较早期脚本串联方式更好。

主要可优化点：
- 步骤间通过文件传递数据，耦合仍偏高（建议中期演进为 Python 函数化调用 + 统一上下文对象）；
- 目前 `051 Send an email.py` 依赖图片文件，但主流程没有显式“生成图片”步骤，建议补充可选步骤并在校验层声明其依赖。

### 中期优化（已开始）

本轮已落地两项结构优化：

1. **减少步骤间文件耦合（契约化）**
   - `pipeline/steps.py` 为每个步骤声明 `input_patterns` / `output_patterns`；
   - `main.py` 在执行前做输入校验，在执行后做输出校验；
   - `pipeline/validators.py` 统一校验逻辑，减少脚本内散落的隐式依赖。

2. **补齐图片生成步骤依赖声明**
   - 主流程新增 `050 image.py`（在发信前生成 `*美的*.png`）；
   - `051 Send an email.py` 明确依赖 `output.html`、`*美的*.png`、`总库存*.xlsx`。

### PyCharm 新手调试指南（单步 / 串联）

你可以按“先单步、再串联”的方式验证：

#### A. 先做三条基础命令（推荐）

```bash
python main.py --check
python main.py --dry-run
python main.py --list-steps
```

- `--check`：看脚本和环境变量是否齐全；
- `--dry-run`：只看执行顺序，不真的处理数据；
- `--list-steps`：确认步骤列表和编号。

#### B. 在 PyCharm 单独运行某个脚本

以 `030 Warehousing at home.py` 为例：

1. 打开 `Run -> Edit Configurations`；
2. 新建 `Python` 配置；
3. `Script path` 选目标脚本（如 `script/030 Warehousing at home.py`）；
4. `Parameters` 填数据目录（如 `data-local`）；
5. `Working directory` 设为仓库根目录；
6. 运行并看控制台输出。

> 单步运行适合排查“某一步逻辑是否正确”。

#### C. 在 PyCharm 跑整条流水线（最常用）

1. 新建 `Python` 配置；
2. `Script path` 设为 `main.py`；
3. `Parameters` 可先填：`--dry-run`；
4. 验证通过后再改为：
   `--data-dir data-local --stop-on-error --report-file data-local/run-report.json`

> 串联运行适合验证“上下游是否衔接正确”。

#### D. 我到底该“单独运行”还是“串联运行”？

- **改了某个业务脚本（如 041/042）**：先单独运行该步骤，再跑整链路；
- **改了主调度（main.py / steps.py / validators.py）**：直接跑 `--dry-run` + 整链路；
- **不确定哪里错**：先 `--check`，再按 `--only-step` 从出错点前一两步开始回放。

#### E. 常见新手坑

- 忘记配置环境变量（邮箱账号/授权码/收件人）；
- Working directory 不是仓库根目录，导致找不到 `script/` 或 `data/`；
- 直接跑全流程，不先 `--dry-run`，排错成本高。

### 最新评分与优化路线

如需查看更详细的“分项评分 + P0/P1/P2 优化优先级”，请看：

- `docs/ARCHITECTURE_SCORECARD.md`

### P0 已落地：020 重试机制 + check 报告细化

1. **020 重试机制（网络敏感步骤）**
   - 默认对 `020 Email download.py` 启用重试（总尝试 = 1 + `--retry-count`，默认重试 2 次）；
   - 退避等待为 `--retry-backoff * attempt` 秒（默认基值 2）。

2. **check 报告细化**
   - `python main.py --check --report-file data/check-report.json` 会输出详细结构化报告；
   - 报告中包含每个步骤的输入/输出契约、脚本存在性、缺失环境变量和重试配置，便于远程定位问题。

> 关于“现在是否还是高耦合”的判断与依据，已补充在 `docs/ARCHITECTURE_SCORECARD.md` 的“是否仍然高耦合？”章节。

> 关于“关键步骤函数化 + manifest 是否有必要”的结论与分阶段策略，见 `docs/ARCHITECTURE_SCORECARD.md` 新增章节。

### 一次性重构（函数化执行）说明

本次已将关键步骤改为可进程内函数执行（默认：`050 image.py`、`050 mailtxt.py`、`051 Send an email.py`）：

- 默认参数：`--in-process-steps "050 image.py,050 mailtxt.py,051 Send an email.py"`
- 仍保留子进程回退能力（不在列表中的步骤继续按原子进程执行）

这让你在 PyCharm 中对关键步骤断点调试更直接，也减少跨进程日志噪音。

> 修复说明：进程内执行步骤时会显式注入 `sys.argv=[script_path, data_dir]`，避免脚本误读取主程序参数（如 `--data-dir`）导致路径解析错误。

> GitHub Actions 修复：`050 image.py` 现在会忽略 `--data-dir` 这类参数标记，自动提取最后一个有效路径参数，避免把 `--data-dir` 误当成目录名。

> 邮件图片样式优化：`050 image.py` 已改为“美化版表格渲染”（分层表头 + 斑马纹 + 轻网格 + 原始高亮色覆盖），优先保证邮件端可读性。

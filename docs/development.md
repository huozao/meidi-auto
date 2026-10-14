# 本地开发与调试

当前业务口径与运行证据见 [CURRENT_HANDOFF.md](CURRENT_HANDOFF.md)，步骤拓扑见
[PIPELINE_FLOW.md](PIPELINE_FLOW.md)。旧教材、架构评分和迁移 FAQ 完整保存在
[历史 README](migrations/2026-09-13-readme-history.md)，不作为当前行为依据。

## 先做无业务副作用的检查

在仓库根目录、已安装 `requirements.txt` 的 Python 环境中运行：

```bash
python main.py --list-steps
python main.py --dry-run
python main.py --check
python tools/generate_module_catalog.py --check
git diff --check
```

`--dry-run` 只列步骤；`--check` 只核对脚本与环境变量，缺变量会告警但不一定失败。
不要把它们当作凭据、远程网络或业务数据验证。文档改动无需真实发信。
代码回归使用 `python -m unittest discover -s tests -v`；现有测试重点覆盖月度汇总，
通过这些测试不代表整条邮件流水线经过真实验证。

## 定位代码

| 修改目标 | 代码入口 |
|---|---|
| 参数、执行顺序筛选、失败策略、运行报告 | `main.py` |
| 生产步骤与输入输出声明 | `pipeline/steps.py`、`pipeline/models.py` |
| 步骤前后校验、共享路径处理 | `pipeline/validators.py`、`pipeline/io_utils.py` |
| 单步计算、邮件、归档 | `script/` 中对应步骤与顶部 STEP CARD |
| 独立月度汇总 | `tools/monthly_summary.py`、[月度契约](MONTHLY_SUMMARY_SPEC.md) |

PyCharm 的解释器选择项目虚拟环境，Working directory 设仓库根目录；
调试 `main.py` 时先用 `--dry-run`，确认参数与步骤。不要通过注释步骤改变执行顺序。
进程内步骤名单由 `main.py` 的 `DEFAULT_IN_PROCESS_STEPS` 定义，非名单步骤使用子进程；
对目标步骤打断点时先确认它采用哪种运行方式。

## 需要真实数据时

`--only-step 030 --data-dir data-local` 会真实处理目录里的数据；选中下载、发信或
归档步骤还会访问真实账户。整链路、本地 Docker 和 GitHub Actions 都没有自动沙箱保证。
获得对应授权后先检查输入和收件人，使用独立目录保留报告，避免并发重复发信。
清理参数会删除产物，应单独核对目标；目录名称含 `local` 不代表可以删除。

`--report-file data-local/run-report.json` 可保存运行结果；
`--check --report-file data-local/check-report.json` 会额外写本地前置检查报告。
检查结果时分清缺环境变量、缺脚本、缺输入和业务失败；原始附件与报告不提交 Git。

## 文档维护

业务行为或流程变化时更新受影响的 STEP CARD、步骤契约、流程图和验收说明。
模块表使用 `docs/module_catalog.json` 生成；修改源表后运行：

```bash
python tools/generate_module_catalog.py
python tools/generate_module_catalog.py --check
```

该检查只验证 Markdown/CSV 与 JSON 一致，不证明模块说明符合源码。
旧 [维护地图](MAINTENANCE_MAP.md)、[架构讨论](AI_FRIENDLY_REDESIGN.md) 与
[评分记录](ARCHITECTURE_SCORECARD.md) 含历史建议；当前路径、输入和实现须对照源码，
不能从旧评分推断待办已完成或必须执行某次重构。

Git 冲突按两侧业务意图逐段合并，不能统一选择 current/incoming/both；
验证后只在用户已授权提交或推送时执行相应操作。

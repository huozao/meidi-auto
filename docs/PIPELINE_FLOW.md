# 流程图（给维护者的 1 分钟导览）

> 目的：降低“步骤变多后不易理解”的阅读成本。

## 总览（Mermaid）

```mermaid
flowchart LR
  S020[020 邮件下载] --> S021[021 合并主文件]
  S021 --> S030[030 家里库存回填]
  S030 --> S032[032 出入库汇总]
  S032 --> S033[033 需求回填]
  S033 --> S041[041 缺口计算]
  S041 --> S042[042 颜色标记]
  S042 --> S050A[050A 生成图片]
  S042 --> S050B[050B 生成HTML]
  S050A --> S051[051 发送邮件]
  S050B --> S051
  S051 --> S052[052 每日附件归档（本地/WebDAV）]
  S051 --> CLEAN[010 清理(可选)]
```

## 各步骤阅读顺序

1. 先看每个脚本顶部 `STEP CARD`（功能/输入/输出/上下游）。
2. 再看 `pipeline/steps.py` 的契约（input_patterns / output_patterns）。
3. 最后看脚本内部业务细节。

## 设计约定

- 每新增或修改主流程脚本，请同步更新脚本顶部 `STEP CARD`。
- 流程拓扑变化（新增步骤/换顺序）时，同步更新本文 Mermaid 图。

## 日报邮件展示约定

- `050 image.py` 使用 LibreOffice Calc 按 `库存表` 的 Excel 样式和打印区域生成 PNG；原始工作簿只读，渲染用副本隐藏其他工作表，默认区域为 `A:T`。不主动改变 Excel 列宽，裁掉页面白边后将图片宽度限制为 1800px，避免邮件附件过宽。
- GitHub runner 需要 `libreoffice-calc`、`fonts-noto-cjk` 和 `poppler-utils`；这些依赖由 `run-daily.yml` 安装。
- `051 Send an email.py` 的标题从 `库存表!H3` 读取业务日期，格式为 `美的库存及出入库日报｜YYYY-MM-DD｜库存、出入库及月计划`，不再使用图片文件名。

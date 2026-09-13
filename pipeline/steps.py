from __future__ import annotations

"""流水线步骤定义（按业务阶段组织）。"""

from .models import PipelineStep

# 阶段 1：邮件数据获取
EMAIL_INGESTION_STEPS: tuple[PipelineStep, ...] = (
    PipelineStep(
        "020 Email download.py",
        output_patterns=("mail_meta.json", "存量查询*.xlsx"),
    ),
)

# 阶段 2：库存主表整形
INVENTORY_TRANSFORMATION_STEPS: tuple[PipelineStep, ...] = (
    PipelineStep(
        "021 Merge excel.py",
        input_patterns=("存量查询*.xlsx",),
        output_patterns=("总库存*.xlsx",),
    ),
    PipelineStep(
        "030 Warehousing at home.py",
        input_patterns=("总库存*.xlsx", "mail_meta.json"),
        output_patterns=("总库存*.xlsx",),
    ),
    PipelineStep(
        "032 Warehousing at out.py",
        input_patterns=("总库存*.xlsx",),
        output_patterns=("总库存*.xlsx",),
    ),
    PipelineStep(
        "033 list insertion.py",
        input_patterns=("总库存*.xlsx",),
        output_patterns=("总库存*.xlsx",),
    ),
    PipelineStep(
        "041 operation.py",
        input_patterns=("总库存*.xlsx",),
        output_patterns=("总库存*.xlsx",),
    ),
)

# 阶段 3：展示内容生成
PRESENTATION_STEPS: tuple[PipelineStep, ...] = (
    PipelineStep(
        "042 Color display.py",
        input_patterns=("总库存*.xlsx",),
        output_patterns=("总库存*.xlsx",),
    ),
    PipelineStep(
        "050 image.py",
        input_patterns=("总库存*.xlsx",),
        output_patterns=("*美的*.png",),
    ),
    PipelineStep(
        "050 mailtxt.py",
        input_patterns=("总库存*.xlsx",),
        output_patterns=("output.html",),
    ),
)

# 阶段 4：邮件通知
NOTIFICATION_STEPS: tuple[PipelineStep, ...] = (
    PipelineStep(
        "051 Send an email.py",
        input_patterns=("output.html", "*美的*.png", "总库存*.xlsx"),
    ),
    PipelineStep(
        "052 Archive daily attachment.py",
        input_patterns=("output.html", "总库存*.xlsx"),
        required=False,
    ),
)


PRODUCTION_STEPS: tuple[PipelineStep, ...] = (
    *EMAIL_INGESTION_STEPS,
    *INVENTORY_TRANSFORMATION_STEPS,
    *PRESENTATION_STEPS,
    *NOTIFICATION_STEPS,
)

CLEANUP_STEP = PipelineStep("010 clean.py", required=False)

from __future__ import annotations

from .models import PipelineStep


PRODUCTION_STEPS: tuple[PipelineStep, ...] = (
    PipelineStep(
        "020 Email download.py",
        output_patterns=("存量查询*.xlsx", "mail_meta.json"),
    ),
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
        support_patterns=("list.xlsx",),
    ),
    PipelineStep(
        "041 operation.py",
        input_patterns=("总库存*.xlsx",),
        output_patterns=("总库存*.xlsx",),
    ),
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
    PipelineStep(
        "051 Send an email.py",
        input_patterns=("output.html", "*美的*.png", "总库存*.xlsx"),
    ),
)

CLEANUP_STEP = PipelineStep("010 clean.py", required=False)

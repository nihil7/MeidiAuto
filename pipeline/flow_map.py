from __future__ import annotations

"""业务主流程导航。

这个模块不负责执行，仅提供“需求 -> 修改点”与“阶段 -> 步骤”的稳定映射，
供新人和 AI 快速定位改动入口。
"""

from dataclasses import dataclass


@dataclass(frozen=True)
class FlowStage:
    """业务阶段定义。"""

    stage_name: str
    goal: str
    step_files: tuple[str, ...]
    key_outputs: tuple[str, ...]


FLOW_STAGES: tuple[FlowStage, ...] = (
    FlowStage(
        stage_name="邮件数据获取",
        goal="从邮箱下载库存源文件并生成基础元数据。",
        step_files=("020 Email download.py",),
        key_outputs=("mail_meta.json", "存量查询*.xlsx"),
    ),
    FlowStage(
        stage_name="库存主表整形",
        goal="将多个来源表合并并加工为单一总库存主表。",
        step_files=(
            "021 Merge excel.py",
            "030 Warehousing at home.py",
            "032 Warehousing at out.py",
            "033 list insertion.py",
            "041 operation.py",
        ),
        key_outputs=("总库存*.xlsx",),
    ),
    FlowStage(
        stage_name="展示内容生成",
        goal="从总库存生成邮件展示素材（Excel 样式、图片、HTML）。",
        step_files=("042 Color display.py", "050 image.py", "050 mailtxt.py"),
        key_outputs=("总库存*.xlsx", "*美的*.png", "output.html"),
    ),
    FlowStage(
        stage_name="通知发送",
        goal="向收件人发送最终库存邮件。",
        step_files=("051 Send an email.py",),
        key_outputs=(),
    ),
)


CHANGE_REQUEST_TO_TARGETS: dict[str, tuple[str, ...]] = {
    "改下载邮件": ("script/020 Email download.py",),
    "改合并库存": ("script/021 Merge excel.py",),
    "改库存加工": (
        "script/030 Warehousing at home.py",
        "script/032 Warehousing at out.py",
        "script/033 list insertion.py",
        "script/041 operation.py",
    ),
    "改展示图片": ("script/042 Color display.py", "script/050 image.py"),
    "改邮件正文": ("script/050 mailtxt.py",),
    "改发信行为": ("script/051 Send an email.py",),
    "改流程顺序": ("pipeline/steps.py",),
    "改校验规则": ("pipeline/validators.py",),
    "改运行策略": ("main.py",),
}

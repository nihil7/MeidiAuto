# AI 友好型项目重构方案（meidi-auto）

> 目标：让新人和 AI 在 **5 分钟内**知道“主流程在哪、改哪里、不要动哪里”。

## 1) 项目目标概述

meidi-auto 的业务主线是：
**从邮件拿库存原始数据 → 合并与清洗 → 生成展示内容（Excel/图片/HTML）→ 发送邮件通知**。

本次重构不追求“层数多”，而追求三件事：

1. **主流程显式可见**：一个文件就能看懂输入、处理、输出。
2. **修改入口显式可见**：需求到代码路径可以一跳定位。
3. **边界显式可见**：哪些是可配置项，哪些必须改代码。

---

## 2) 推荐目录结构（按业务能力分区）

```text
meidi-auto/
├─ main.py                          # 统一入口：参数解析 + 运行主流程
├─ pipeline/
│  ├─ steps.py                      # 业务步骤清单（按业务阶段组织）
│  ├─ flow_map.py                   # 主流程导航与“需求 -> 修改点”映射
│  ├─ validators.py                 # 关键输入/输出校验（防级联误伤）
│  ├─ models.py                     # Step 数据模型（稳定边界）
│  └─ io_utils.py                   # 文件发现与目录解析（仅 I/O 辅助）
├─ script/                          # 具体业务脚本（每步一个明确动作）
│  ├─ 020 Email download.py
│  ├─ 021 Merge excel.py
│  ├─ 030 Warehousing at home.py
│  ├─ 032 Warehousing at out.py
│  ├─ 033 list insertion.py
│  ├─ 041 operation.py
│  ├─ 042 Color display.py
│  ├─ 050 image.py
│  ├─ 050 mailtxt.py
│  └─ 051 Send an email.py
├─ docs/
│  ├─ PIPELINE_FLOW.md              # 流程图
│  ├─ MAINTENANCE_MAP.md            # 维护导航
│  └─ AI_FRIENDLY_REDESIGN.md       # 本文档：AI/新人导向架构说明
└─ tools/
   └─ generate_module_catalog.py    # 文档索引生成工具（非业务主线）
```

> 说明：保留现有可运行结构，重点把“主线、导航、边界”显式化，避免大拆大改导致生产风险。

---

## 3) 每个目录/模块的一句话职责说明

- `main.py`：只做“流程编排与执行控制”，不承载业务细节。
- `pipeline/steps.py`：定义“做什么”，按业务阶段列出所有步骤。
- `pipeline/flow_map.py`：回答“某个需求改哪里、影响哪些步骤”。
- `pipeline/validators.py`：回答“某步能不能执行、执行后产物是否达标”。
- `script/`：回答“这一步具体怎么做”。
- `docs/`：回答“新人和 AI 如何快速理解与修改”。

---

## 4) 主流程说明：输入 -> 处理 -> 输出

### 输入

- 邮件系统中的库存附件（020 下载）；
- 环境变量（邮箱账号、密码、收件人等）；
- 可选命令参数（`--data-dir`、`--only-step` 等）。

### 处理

1. **数据获取阶段**：`020` 下载邮件与附件。
2. **数据整形阶段**：`021` 合并，`030/032/033/041` 逐步加工库存主表。
3. **展示产物阶段**：`042` 着色，`050 image` 生成图片，`050 mailtxt` 生成邮件 HTML。
4. **通知阶段**：`051` 发送邮件。
5. **可选清理阶段**：`010 clean` 清理产物。

### 输出

- `总库存*.xlsx`（业务主数据）；
- `*美的*.png`（展示图）；
- `output.html`（邮件正文）；
- 邮件发送结果与运行报告 JSON。

---

## 5) 常见修改需求应该去改哪里

- “改邮件下载逻辑” → `script/020 Email download.py`
- “改库存合并规则” → `script/021 Merge excel.py`
- “改库存加工口径（入库/出库/插入/运营）” → `script/030/032/033/041`
- “改可视化样式（颜色/图片）” → `script/042 Color display.py`、`script/050 image.py`
- “改邮件正文模板” → `script/050 mailtxt.py`
- “改发送策略/收件人处理” → `script/051 Send an email.py`
- “改步骤顺序或新增步骤” → `pipeline/steps.py`
- “改产物校验规则” → `pipeline/validators.py`
- “改运行参数/重试/失败策略” → `main.py`

---

## 6) 哪些部分禁止轻易修改

1. `PipelineStep` 字段语义（`filename/input_patterns/output_patterns`）
   - 这是编排层和脚本层的契约；随意改字段会导致全链路失效。
2. `pipeline/validators.py` 的关键产物校验逻辑
   - 这是防止“子脚本误返回成功”导致级联错误的保险丝。
3. `main.py` 的失败处理与报告结构
   - 外部自动化依赖这些结果做判断。

> 原则：若必须改以上边界，先更新 `docs/MAINTENANCE_MAP.md` 与 `docs/PIPELINE_FLOW.md`，再改代码。

---

## 7) 配置项与代码逻辑的边界划分

### 可配置（优先通过参数/环境变量调整）

- 运行目录：`--data-dir`、`--script-dir`
- 运行策略：`--stop-on-error`、`--retry-count`、`--retry-backoff`
- 执行范围：`--only-step`、`--clean-only`、`--clean-after-run`
- 邮件连接参数：`EMAIL_ADDRESS_QQ`、`EMAIL_PASSWORD_QQ|EMAIL_PASSWOR_QQ`、`RECIPIENT_EMAILS`

### 必须改代码（业务规则发生变化）

- Excel 字段口径、库存计算规则
- 展示层图像生成算法
- 邮件正文业务逻辑
- 步骤间输入/输出契约

---

## 8) 骨架代码与主流程文件建议

当前建议以 `main.py + pipeline/steps.py` 作为双主线：

- `main.py`：运行时主流程（真实执行入口）
- `pipeline/steps.py`：业务动作主线（静态导航入口）

新增 `pipeline/flow_map.py` 用于输出“阶段 -> 步骤 -> 产物 -> 修改入口”的结构化导航，便于 AI 在接到自然语言需求时直接定位修改点。

---

## 9) 对原始设计的不符合点与重构结论

不符合 AI 友好原则的点：

1. **步骤列表此前是单段线性列表**：能跑，但“业务阶段感”不强，AI 很难第一眼判断每步属于哪条业务能力。
2. **缺少统一“需求 -> 修改点”映射文件**：新人要靠搜索猜路径，容易误改。
3. **主流程说明散落在多个文档**：阅读成本高，定位慢。

重构结论：

- 通过 `pipeline/steps.py` 的“按业务阶段分组”强化主线可读性；
- 通过 `pipeline/flow_map.py` 固化修改导航和边界说明；
- 用本文件把“结构理由 + 修改策略 + 禁改边界”一次讲清。

这套组织比“纯技术分层 + 到处跳转”更适合 AI 持续接手，因为它优先回答：
**“现在要改什么、在哪改、改完会影响谁”。**

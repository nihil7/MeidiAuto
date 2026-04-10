# MeidiAuto2 维护地图（模块功能 / 冗余评估 / 维护建议）

## 1) 主流程模块（生产链路）

> 由 `main.py` + `pipeline/steps.py` 串联执行，属于默认生产路径。

| 步骤 | 文件 | 功能 | 输入 | 关键输出 | 备注 |
|---|---|---|---|---|---|
| 020 | `script/020 Email download.py` | 连接 IMAP，抓取邮件 HTML 与附件，产出元数据 | 邮箱环境变量、IMAP | `mail_meta.json`、`存量查询*.xlsx` | 生产入口数据源 |
| 021 | `script/021 Merge excel.py` | 以最新“合肥市”文件为基底，合并库存相关表 | 020 导出的 Excel | `总库存*.xlsx` | 生产主文件 |
| 030 | `script/030 Warehousing at home.py` | 加工库存表结构、家里库存回填、时间写入 | `总库存*.xlsx`、`mail_meta.json` | 更新后的 `总库存*.xlsx` | 强依赖 020 元数据 |
| 032 | `script/032 Warehousing at out.py` | 汇总出入库明细并新建汇总工作表 | `总库存*.xlsx` | “出入库汇总和其他变动”sheet | 侧重出入库统计 |
| 033 | `script/033 list insertion.py` | 从需求表写入 K/N/P/T 列并统一格式 | `script/data/list.xlsx`、`总库存*.xlsx` | 更新后的 `总库存*.xlsx` | 依赖固定需求模板 |
| 041 | `script/041 operation.py` | 计算最小发货/排产/月计划缺口，并写合计 | `总库存*.xlsx` | 更新后的 `总库存*.xlsx` | 列名敏感，改表头需同步 |
| 042 | `script/042 Color display.py` | 按业务规则给库存表着色（深色+淡色扩展） | `总库存*.xlsx` | 更新后的 `总库存*.xlsx` | 规则可配置 |
| 050A | `script/050 image.py` | 生成邮件图片附件（*美的*.png） | `总库存*.xlsx` | `*美的*.png` | 为 051 提供图片依赖 |
| 050B | `script/050 mailtxt.py` | 生成邮件正文 HTML（包含异常行等） | `总库存*.xlsx` | `output.html` | 供 051 发送 |
| 051 | `script/051 Send an email.py` | 读取 HTML / 图片 / Excel 并发送邮件 | `output.html`、`*美的*.png`、`总库存*.xlsx` | 外发邮件 | 终端动作 |
| clean | `script/010 clean.py` | 清理中间产物（支持 data/script/data） | 数据目录 | 删除冗余文件 | 可独立执行 |

---

## 2) 编排与基础设施模块

| 文件 | 功能 | 修改建议 |
|---|---|---|
| `main.py` | 参数解析、步骤调度、失败中断、报告输出、dry-run/check | 新增流程优先改这里，不建议再写并行“本地主程序” |
| `pipeline/steps.py` | 生产步骤清单（执行顺序） | 增删步骤的单一入口 |
| `pipeline/validators.py` | 必要脚本存在校验、020/021关键产物校验 | 新增关键产物时补充校验规则 |
| `pipeline/io_utils.py` | 脚本共享：解析数据目录、查找 Excel | 统一文件发现逻辑，避免每个脚本重复写 glob |
| `pipeline/models.py` | `PipelineStep` 数据结构 | 基础模型，通常稳定 |

---

## 3) 冗余脚本处理结果（已执行删除）

以下历史/弱关联脚本已从仓库删除，以降低维护噪音：

- `script/main local.py`
- `script/企业消息整理.py`
- `script/050 image local.py`
- `script/月汇总.py`
- `script/月汇总绘图.py`
- `script/统一格式.py`
- `script/sync_env_to_github.py`
- `多网站.py`
- `chek.py`

删除依据：与主流程重复、平台依赖强、或与库存主业务弱关联。

---

## 4) 推荐维护方式（让修改者更快定位）

### A. 目录分层（优先级最高）
- `script/` 只放生产步骤（020~051 + 010 clean）。
- `script/legacy/` 放历史脚本（保留但不默认执行）。
- `tools/` 放运维/一次性工具（如同步 secrets、本地格式化辅助等）。

### B. “改哪里”速查
- 改执行顺序/是否运行某步骤：`pipeline/steps.py`
- 改调度参数/失败策略/report 输出：`main.py`
- 改“某步骤跑完是否算成功”：`pipeline/validators.py`
- 改通用找文件/路径策略：`pipeline/io_utils.py`
- 改具体业务逻辑：对应 `script/0xx ...py`

### C. 变更流程建议（最少回归路径）
1. `python main.py --check`
2. `python main.py --dry-run`
3. 单步验证（例如 `--only-step 030`）
4. 全量执行并产出报告（`--report-file`）

### D. 冗余治理策略
- 第一阶段：**标记但不删除**（改名为 `*.legacy.py` 或移目录）。
- 第二阶段：连续 2~4 周无人使用再删除。
- 删除前保留“脚本迁移对照表”（旧文件 -> 新入口）。

---

## 5) 给后续维护者的结论

- 当前仓库已经有统一调度入口（`main.py`），这是后续维护的“唯一主入口”。
- 最大维护成本来自“主流程脚本 + 历史脚本并存”，建议尽快做目录分层和冗余归档。
- 修改时优先遵循“先改编排层、后改步骤脚本、最后补校验”的顺序，可显著降低回归风险。

---

## 6) 表格维护与生成功能（新增）

为减少手工维护成本，现已提供“模块清单自动生成功能”：

- 编辑 `docs/module_catalog.json`（单一真源）
- 运行 `python tools/generate_module_catalog.py` 自动生成：
  - `docs/MODULE_CATALOG.md`
  - `docs/module_catalog.csv`
- 运行 `python tools/generate_module_catalog.py --check` 可校验生成文件是否过期（适合 CI）。

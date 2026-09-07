# 模块清单（自动生成）

本文件由 `tools/generate_module_catalog.py` 自动生成，请勿手工修改。

## 主流程模块

| Step | 文件 | 分类 | 功能 | 输入 | 输出 |
|---|---|---|---|---|---|
| 020 | `script/020 Email download.py` | 邮件采集 | 下载邮件HTML与附件并生成mail_meta元数据 | EMAIL_*环境变量, IMAP_SERVER | mail_meta.json, 存量查询*.xlsx |
| 021 | `script/021 Merge excel.py` | 数据整合 | 合并库存Excel并生成总库存主文件 | 020阶段Excel | 总库存*.xlsx |
| 030 | `script/030 Warehousing at home.py` | 库存加工 | 加工库存表结构并回填家里库存 | 总库存*.xlsx, mail_meta.json | 更新后的总库存*.xlsx |
| 032 | `script/032 Warehousing at out.py` | 库存加工 | 汇总出入库明细并生成汇总工作表 | 总库存*.xlsx | 更新后的总库存*.xlsx |
| 033 | `script/033 list insertion.py` | 计划回填 | 将需求表数据写入库存关键列并统一格式 | script/data/list.xlsx, 总库存*.xlsx | 更新后的总库存*.xlsx |
| 041 | `script/041 operation.py` | 计算汇总 | 计算最小发货/排产/月计划缺口与合计 | 总库存*.xlsx | 更新后的总库存*.xlsx |
| 042 | `script/042 Color display.py` | 可视化标记 | 按业务阈值对库存表着色 | 总库存*.xlsx | 更新后的总库存*.xlsx |
| 050A | `script/050 image.py` | 消息素材 | 将库存表导出为邮件图片附件 | 总库存*.xlsx | *美的*.png |
| 050B | `script/050 mailtxt.py` | 消息生成 | 生成邮件HTML正文 | 总库存*.xlsx | output.html |
| 051 | `script/051 Send an email.py` | 消息发送 | 发送邮件并附带HTML/图片/Excel | output.html, *美的*.png, 总库存*.xlsx | 邮件发送结果 |
| clean | `script/010 clean.py` | 清理 | 清理中间产物与冗余文件 | data目录 | 清理结果日志 |

## 历史/工具模块

| 文件 | 状态 | 说明 |
|---|---|---|

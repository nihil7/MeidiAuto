import os
import sys
import glob
import openpyxl
from collections import defaultdict

# ================================
# 📂 文件路径配置（支持主程序传参）
# ================================
default_inventory_folder = r'C:\Users\ishel\Desktop\当日库存情况'

# 通过 sys.argv 传递路径参数
if len(sys.argv) >= 2:
    inventory_folder = sys.argv[1]
    print(f"✅ 使用传入路径: {inventory_folder}")
else:
    inventory_folder = default_inventory_folder
    print(f"⚠️ 未传入路径，使用默认路径: {inventory_folder}")

# 判断路径是否存在
if not os.path.exists(inventory_folder):
    print(f"❌ 文件夹路径不存在: {inventory_folder}")
    sys.exit(1)

print(f"📂 当前工作文件夹: {inventory_folder}")

# 匹配文件：总库存*.xlsx
pattern = os.path.join(inventory_folder, '总库存*.xlsx')
files = glob.glob(pattern)

# 判断文件是否找到
if not files:
    print("❌ 没有找到符合条件的文件！")
    exit()

inventory_file = files[0]  # 取第一个文件
print(f"✅ 找到文件：{inventory_file}")

# ================================
# 2. 打开Excel文件，读取工作表
# ================================
wb_inventory = openpyxl.load_workbook(inventory_file)

# 检查工作表是否存在
sheet_name_detail = '出入库明细表'
if sheet_name_detail not in wb_inventory.sheetnames:
    print(f"❌ 没有找到工作表：{sheet_name_detail}")
    exit()

sheet_detail = wb_inventory[sheet_name_detail]

# ================================
# 3. 读取表头和列索引
# ================================
# 表头在第3行
header_row_index = 3

headers = [cell.value for cell in sheet_detail[header_row_index]]
print(f"✅ 表头内容：{headers}")

# 获取列名和索引（列索引从1开始）
col_idx = {header: idx + 1 for idx, header in enumerate(headers)}

# ✅ 修改这里，改为你表头里真实存在的列名
required_columns = ['库存变动类别', '美的编码', '本期收入', '本期发出', '出入库日期']

# 检查必要列是否存在
for col in required_columns:
    if col not in col_idx:
        print(f"❌ 缺少必要列：{col}")
        print(f"🔎 当前列索引：{col_idx}")
        exit()

# ================================
# 4. 分类和汇总数据
# ================================
summary_data = defaultdict(lambda: {'入库': 0, '出库': 0})
other_records = []

# 数据从表头下一行开始读取（第4行）
for row in sheet_detail.iter_rows(min_row=header_row_index + 1, values_only=True):
    变动类别 = row[col_idx['库存变动类别'] - 1]
    美的编码 = row[col_idx['美的编码'] - 1]
    本期收入 = row[col_idx['本期收入'] - 1] or 0
    本期发出 = row[col_idx['本期发出'] - 1] or 0
    出入库日期 = row[col_idx['出入库日期'] - 1]

    if 变动类别 == '入库':
        summary_data[美的编码]['入库'] += 本期收入
    elif 变动类别 == '出库':
        summary_data[美的编码]['出库'] += 本期发出
    else:
        other_records.append(row)

# ================================
# 5. 创建/更新工作表
# ================================

# 新建一个工作表用于“出入库汇总”和“其他变动明细”
sheet_name_combined = '出入库汇总和其他变动'
if sheet_name_combined in wb_inventory.sheetnames:
    del wb_inventory[sheet_name_combined]
sheet_combined = wb_inventory.create_sheet(sheet_name_combined)

# ---- 出入库汇总 ----
# 写表头
sheet_combined.append(['美的编码', '本期收入（入库）', '本期发出（出库）'])

# 写入数据
for 编码, data in summary_data.items():
    sheet_combined.append([编码, data['入库'], data['出库']])

# 获取“出入库汇总”表的最后一行
last_row_summary = len(sheet_combined['A'])

# ---- 其他变动明细 ----
# 计算当前表中连续空行的数量
empty_rows = 0
for row in sheet_combined.iter_rows(min_row=last_row_summary + 1, max_row=last_row_summary + 5, values_only=True):
    if all(cell is None for cell in row):
        empty_rows += 1
    else:
        break

# 控制空行不超过5行，如果超过，则补充最多5行空行
if empty_rows < 5:
    for _ in range(5 - empty_rows):
        sheet_combined.append([None] * len(headers))  # 添加空行以确保5行空行

# 写表头
sheet_combined.append(['录入日期', '客户子库', '单号', '美的编码', '物料品名', '单位', '仓库', '库存变动类别', '本期收入', '本期发出', '条形码', '备注', '代编码', '出入库日期'])

# 写入其他记录
for record in other_records:
    sheet_combined.append(record)

# ================================
# 6. 自动调整列宽
# ================================
for col in sheet_combined.columns:
    max_length = 0
    column = col[0].column_letter  # 获取列字母
    for cell in col:
        try:
            if len(str(cell.value)) > max_length:
                max_length = len(cell.value)
        except:
            pass
    adjusted_width = (max_length + 2)
    sheet_combined.column_dimensions[column].width = adjusted_width

# ================================
# 7. 将“出入库汇总和其他变动”第一列与“库存表”第二列匹配
# ================================
# 打开库存表工作表
sheet_inventory = wb_inventory['库存表']

# 获取“出入库汇总和其他变动”的数据（第一列和第二列）
summary_first_col = [row[0] for row in sheet_combined.iter_rows(min_row=2, max_row=last_row_summary + 1, values_only=True)]
summary_second_col = [row[1] for row in sheet_combined.iter_rows(min_row=2, max_row=last_row_summary + 1, values_only=True)]

# 获取“库存表”的第二列数据
inventory_second_col = [row[1] for row in sheet_inventory.iter_rows(min_row=2, max_row=sheet_inventory.max_row, values_only=True)]

# 遍历“库存表”第二列，匹配并复制数据到第17列
for idx, inventory_value in enumerate(inventory_second_col):
    if inventory_value in summary_first_col:
        summary_index = summary_first_col.index(inventory_value)
        sheet_inventory.cell(row=idx + 2, column=17, value=summary_second_col[summary_index])

# ================================
# 8. 保存文件
# ================================
# 保存到原文件中
wb_inventory.save(inventory_file)

print(f"✅ 完成！文件已保存到：{inventory_file}")

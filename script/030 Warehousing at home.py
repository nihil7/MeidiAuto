import os
import sys
import re
import glob
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Border, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.styles import NamedStyle

# ================================
# 📂 路径配置（支持主程序传参）
# ================================
default_folder_path = os.path.join(os.getcwd(), "data")  # GitHub 使用相对路径

if len(sys.argv) >= 2:
    folder_path = sys.argv[1]
else:
    folder_path = default_folder_path

if not os.path.exists(folder_path):
    print(f"❌ 文件夹路径不存在: {folder_path}")
    sys.exit(1)

files = glob.glob(os.path.join(folder_path, "*总库存*.xlsx"))
if not files:
    print("没有找到含有'总库存'的Excel文件！")
    sys.exit(1)

latest_file = max(files, key=os.path.getmtime)
merged_wb = load_workbook(latest_file)
if "库存表" not in merged_wb.sheetnames:
    print("⚠️ 找不到 '库存表' 工作表！")
    sys.exit(1)

sheet_kc = merged_wb["库存表"]

# 解除合并单元格（排除第一行）
merged_cells_ranges = list(sheet_kc.merged_cells.ranges)
for merged_range in merged_cells_ranges:
    if merged_range.min_row > 1:
        sheet_kc.unmerge_cells(str(merged_range))

# 插入列
sheet_kc.insert_cols(10, 10)
sheet_kc.insert_cols(3, 1)

# 让第 3 列（C 列）居中
for cell in sheet_kc["C"]:
    cell.alignment = Alignment(horizontal="center", vertical="center")

# 重新合并单元格
sheet_kc.merge_cells('H3:J3')
# 合并 R3:T3
sheet_kc.merge_cells('U3:W3')
# 只在左上角单元格 R3 赋值
sheet_kc["U3"] = "不合格"


# 提取第2列数据的后4位，写入第3列
for row in sheet_kc.iter_rows(min_row=2, max_col=3):
    cell_value = str(row[1].value).strip() if row[1].value else ""
    match = re.search(r'\d+', cell_value)
    if match:
        row[2].value = match.group()[-4:].zfill(4)

sheet_kc["C4"] = "编号"

# 设置 K4:O4 标题及格式
header_titles = ["外应存", "最小发货", "家里库存", "家应存", "排产", "月计划", "月计划缺口", "外仓出库总量", "外仓入库总量"]
for i, title in enumerate(header_titles):
    col_letter = chr(ord('K') + i)
    cell = sheet_kc[f"{col_letter}4"]
    cell.value = title
    cell.fill = PatternFill(start_color="C187F7", end_color="C187F7", fill_type="solid")

# 处理第一页副本
if "第一页" in merged_wb.sheetnames:
    sheet_first = merged_wb["第一页"]
    header_row = [cell.value for cell in sheet_first[1]]
    if "仓库" in header_row:
        warehouse_col_index = header_row.index("仓库") + 1
        data_to_copy = [row for row in sheet_first.iter_rows(min_row=2, values_only=True) if row[warehouse_col_index - 1] == "成品库"]
        sheet_copy = merged_wb.create_sheet("第一页副本")
        sheet_copy.append(header_row)
        for row in data_to_copy:
            sheet_copy.append(row)

# 处理家里库存表
if "第一页副本" in merged_wb.sheetnames:
    sheet_copy = merged_wb["第一页副本"]
    header_row = [cell.value for cell in sheet_copy[1]]

    if "存货名称" in header_row:
        inventory_name_col_index = header_row.index("存货名称") + 1
        extracted_data = []

        for row in sheet_copy.iter_rows(min_row=2, values_only=True):
            # 处理存货名称
            item_name = str(row[inventory_name_col_index - 1]).strip() if row[inventory_name_col_index - 1] else ""
            four_digits = item_name[:4] if not all('\u4e00' <= char <= '\u9fa5' for char in item_name[:5]) else ""

            # 处理数量列，确保是数值格式
            quantity = row[header_row.index("主数量")] if "主数量" in header_row else None

            # 确保 quantity 为数值，避免存成文本
            if isinstance(quantity, str):
                quantity = float(quantity) if quantity.replace(".", "", 1).isdigit() else None

            extracted_data.append([four_digits, item_name, quantity])

        # 创建 "家里库存" 表
        home_stock_sheet = merged_wb.create_sheet("家里库存")
        home_stock_sheet.append(["编号", "存货名称", "数量"])

        for data in extracted_data:
            home_stock_sheet.append(data)

        # 让 Excel 识别“数量”列为数值格式
        for row in home_stock_sheet.iter_rows(min_row=2, max_row=home_stock_sheet.max_row, min_col=3, max_col=3):
            for cell in row:
                if cell.value is not None:
                    cell.number_format = "#,##0.00"  # 应用千位分隔格式

# 进行库存比对
if "家里库存" in merged_wb.sheetnames:
    home_stock_sheet = merged_wb["家里库存"]
    home_stock_dict = {str(row[0]).zfill(4): row[2] for row in home_stock_sheet.iter_rows(min_row=2, values_only=True)}
    for row in sheet_kc.iter_rows(min_row=2, max_col=13):
        third_col_value = str(row[2].value).strip() if row[2].value else ""
        if third_col_value in home_stock_dict:
            row[12].value = home_stock_dict[third_col_value]

# 2️⃣ 找到 B 列（第 2 列）最下面的 **第一个空单元格所在行号**
col_B = sheet_kc["B"]  # 选取 B 列
max_row = sheet_kc.max_row  # 获取 Excel 认为的最大行数
last_empty_row = max_row + 1  # 默认值，如果 B 列没有空行，则返回 max_row+1

for row in range(4, max_row + 1):  # 从第 3 行开始正向遍历
    if sheet_kc[f"B{row}"].value is None:
        last_empty_row = row
        break  # 只找第一个空单元格

print(f"⚡ 发现 B 列第一个空单元格所在行: {last_empty_row}")


if last_empty_row is None:  # 兜底处理，如果 B 列没有空位，则放在最后一行+1
    last_empty_row = sheet_kc.max_row + 1

print(f"✅ 计算求和的目标行: {last_empty_row}")  # 先检查行号是否符合预期

# 3️⃣ 计算 G 到 U 列（7~21列）的各列总和，并填入 last_empty_row 行
for col in range(7, 22):  # G 到 U 对应列 7~21
    col_letter = get_column_letter(col)  # 获取列字母

    if col == 12:  # 仅针对 L 列（第 12 列）
        sum_formula = f"=SUMIF({col_letter}5:{col_letter}{last_empty_row - 1}, \">0\")"
    else:
        sum_formula = f"=SUM({col_letter}5:{col_letter}{last_empty_row - 1})"

    sum_cell = sheet_kc[f"{col_letter}{last_empty_row}"]  # 目标单元格
    sum_cell.value = sum_formula  # 填入求和公式

# 5️⃣ 右对齐 + 直接应用会计格式（整个数据列 + 求和行）
for col in range(7, 22):  # G 到 Q（7~21）
    col_letter = get_column_letter(col)
    format_range = f"{col_letter}5:{col_letter}{last_empty_row}"  # 从第5行到求和行

    for cell in sheet_kc[format_range]:  # 遍历列的所有单元格
        for c in cell:  # cell 是一个 tuple，遍历其中的 Cell 对象
            c.alignment = Alignment(horizontal="right")  # ✅ 右对齐
            c.number_format = '_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)'  # ✅ 会计格式（千分位）



# 自动调整列宽
from openpyxl.utils import get_column_letter

def auto_adjust_column_width(sheet):
    column_widths = {}
    for row in sheet.iter_rows():
        for cell in row:
            if cell.value and not isinstance(cell, type(None)) and not isinstance(cell, MergedCell):
                col_letter = get_column_letter(cell.column)
                column_widths[col_letter] = max(column_widths.get(col_letter, 0), len(str(cell.value)))

    for col_letter, width in column_widths.items():
        sheet.column_dimensions[col_letter].width = width + 2  # 适当增加宽度



# 保存文件
merged_wb.save(latest_file)
merged_wb.close()

print("🎉 全部处理完成！")
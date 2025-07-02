import os
import sys
import re
import glob
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Alignment
from openpyxl.utils import get_column_letter

# ================================
# 📂 1️⃣ 路径配置（支持主程序传参）
# ================================
default_folder_path = os.path.join(os.getcwd(), "data")  # 如果未传参，使用当前工作目录下 data 文件夹

if len(sys.argv) >= 2:
    folder_path = sys.argv[1]
else:
    folder_path = default_folder_path

if not os.path.exists(folder_path):
    print(f"❌ 文件夹路径不存在: {folder_path}")
    sys.exit(1)

# 查找最新的 *总库存*.xlsx 文件
files = glob.glob(os.path.join(folder_path, "*总库存*.xlsx"))
if not files:
    print("没有找到含有'总库存'的Excel文件！")
    sys.exit(1)

latest_file = max(files, key=os.path.getmtime)  # 获取最新文件
merged_wb = load_workbook(latest_file)

if "库存表" not in merged_wb.sheetnames:
    print("⚠️ 找不到 '库存表' 工作表！")
    sys.exit(1)

sheet_kc = merged_wb["库存表"]

# ================================
# 📊 2️⃣ 找到 B 列第一个空单元格所在行（作为有效数据范围）
# ================================
max_row = sheet_kc.max_row
last_empty_row = max_row + 1  # 如果找不到空行，使用 max_row + 1

for row in range(4, max_row + 1):
    if sheet_kc[f"B{row}"].value is None:
        last_empty_row = row
        break

print(f"⚡ 发现 B 列第一个空单元格所在行: {last_empty_row}")

# ================================
# 🪓 3️⃣ 解除所有合并单元格（包括第一行）
# ================================
for merged_range in list(sheet_kc.merged_cells.ranges):
    sheet_kc.unmerge_cells(str(merged_range))

# ================================
# ➕ 4️⃣ 插入列以准备后续数据填写
# ================================
sheet_kc.insert_cols(10, 10)  # 在 J 列后插入 10 列
sheet_kc.insert_cols(3, 1)    # 在 B 列后插入 1 列（用于存放提取的编号）

# ================================
# 🎯 5️⃣ 设置 C 列（编号列）居中，设置 B1 左对齐
# ================================
for cell in sheet_kc["C"]:
    cell.alignment = Alignment(horizontal="center", vertical="center")

sheet_kc["B1"].alignment = Alignment(horizontal="left", vertical="center")

# ================================
# 🪄 6️⃣ 合并标题单元格，并设置“不合格”标题
# ================================
sheet_kc.merge_cells('H3:J3')
sheet_kc.merge_cells('U3:W3')
sheet_kc["U3"] = "不合格"

# ================================
# 🆔 7️⃣ 提取第二列编号（后5位）写入第三列（编号列）
# ================================
for row in sheet_kc.iter_rows(min_row=2, max_row=last_empty_row - 1, max_col=3):
    cell_value = str(row[1].value).strip() if row[1].value else ""
    match = re.search(r'\d+', cell_value)
    if match:
        row[2].value = match.group()[-5:].zfill(5)  # 保留后5位并补足0
sheet_kc["C4"] = "编号"

# ================================
# 🪄 8️⃣ 设置 K4:O4 表头内容及填充底色
# ================================
header_titles = ["外应存", "最小发货", "家里库存", "家应存", "排产", "月计划", "月计划缺口", "外仓出库总量", "外仓入库总量"]
for i, title in enumerate(header_titles):
    col_letter = chr(ord('K') + i)
    cell = sheet_kc[f"{col_letter}4"]
    cell.value = title
    cell.fill = PatternFill(start_color="C187F7", end_color="C187F7", fill_type="solid")

# ================================
# 📄 9️⃣ 创建“第一页副本”工作表（仅保留仓库=成品库）
# ================================
if "第一页" in merged_wb.sheetnames:
    sheet_first = merged_wb["第一页"]
    header_row = [cell.value for cell in sheet_first[1]]

    if "仓库" in header_row:
        warehouse_col_index = header_row.index("仓库") + 1
        data_to_copy = [
            row for row in sheet_first.iter_rows(min_row=2, values_only=True)
            if row[warehouse_col_index - 1] == "成品库"
        ]
        sheet_copy = merged_wb.create_sheet("第一页副本")
        sheet_copy.append(header_row)
        for row in data_to_copy:
            sheet_copy.append(row)

# ================================
# 📄 10️⃣ 创建“家里库存”表，提取编号、存货名称、数量
# ================================
if "第一页副本" in merged_wb.sheetnames:
    sheet_copy = merged_wb["第一页副本"]
    header_row = [cell.value for cell in sheet_copy[1]]

    if "存货名称" in header_row:
        inventory_name_col_index = header_row.index("存货名称") + 1
        extracted_data = []

        for row in sheet_copy.iter_rows(min_row=2, values_only=True):
            item_name = str(row[inventory_name_col_index - 1]).strip() if row[inventory_name_col_index - 1] else ""
            five_digits = item_name[:5] if not all('\u4e00' <= char <= '\u9fa5' for char in item_name[:5]) else ""
            quantity = row[header_row.index("主数量")] if "主数量" in header_row else None

            if isinstance(quantity, str):
                quantity = float(quantity) if quantity.replace(".", "", 1).isdigit() else None

            extracted_data.append([five_digits, item_name, quantity])

        home_stock_sheet = merged_wb.create_sheet("家里库存")
        home_stock_sheet.append(["编号", "存货名称", "数量"])
        for data in extracted_data:
            home_stock_sheet.append(data)

        # 千位分隔格式设置
        for row in home_stock_sheet.iter_rows(min_row=2, max_row=home_stock_sheet.max_row, min_col=3, max_col=3):
            for cell in row:
                if cell.value is not None:
                    cell.number_format = "#,##0.00"

# ================================
# 🔄 11️⃣ 比对“家里库存”与“库存表”，写入数量到 M 列
# ================================
if "家里库存" in merged_wb.sheetnames:
    home_stock_sheet = merged_wb["家里库存"]

    # 后 4 位 → 行映射
    inventory_suffix_dict = {
        str(row[2].value)[-4:]: row for row in sheet_kc.iter_rows(min_row=2, max_row=last_empty_row - 1, max_col=13)
        if row[2].value
    }
    # 5 位标准编号 → 行映射
    inventory_code_dict = {
        str(row[2].value).zfill(5): row for row in sheet_kc.iter_rows(min_row=2, max_row=last_empty_row - 1, max_col=13)
        if row[2].value
    }

    for row in home_stock_sheet.iter_rows(min_row=2, values_only=True):
        raw_code = str(row[0]).strip() if row[0] else ""
        quantity = row[2]

        if re.fullmatch(r"\d{4}-", raw_code):
            key4 = raw_code[:4]
            if key4 in inventory_suffix_dict:
                inventory_suffix_dict[key4][12].value = quantity
        elif re.fullmatch(r"\d{5}", raw_code):
            key5 = raw_code.zfill(5)
            if key5 in inventory_code_dict:
                inventory_code_dict[key5][12].value = quantity

# ================================
# 🎨 12️⃣ 批量格式设置（G～Q列会计格式，列宽、冻结、缩放）
# ================================
for col in range(7, 18):
    col_letter = get_column_letter(col)
    for row in range(5, last_empty_row + 1):
        cell = sheet_kc[f"{col_letter}{row}"]
        cell.alignment = Alignment(horizontal="right")
        cell.number_format = '_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)'

col_widths = {
    'B': 4.5, 'C': 5, 'D': 35.88, 'E': 3, 'F': 3.6,
    'G': 8.6, 'H': 8, 'I': 8, 'J': 8,
    'K': 5.88, 'L': 8.1, 'M': 9.8, 'N': 5.88,
    'O': 9.5, 'P': 9.8, 'Q': 10.08, 'R': 9.5, 'S': 9.5
}
for col, width in col_widths.items():
    sheet_kc.column_dimensions[col].width = width + 0.6

sheet_kc.row_dimensions[1].height = 18
sheet_kc.freeze_panes = "A5"
sheet_kc.row_dimensions[2].outlineLevel = 1
sheet_kc.row_dimensions[2].hidden = True
sheet_kc.sheet_properties.outlinePr.summaryBelow = True
sheet_kc.sheet_view.zoomScale = 95

print("✅ 格式处理完成")

# ================================
# 💾 13️⃣ 保存并关闭文件
# ================================
merged_wb.save(latest_file)
merged_wb.close()
print(f"🎉 已完成处理并保存: {latest_file}")

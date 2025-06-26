import os
import sys
import re
import glob
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Alignment
from openpyxl.utils import get_column_letter

# ================================
# 📂 路径配置（支持主程序传参）
# ================================
default_folder_path = os.path.join(os.getcwd(), "data")

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

# 2️⃣ 找到 B 列（第 2 列）最下面的 **第一个空单元格所在行号**
col_B = sheet_kc["B"]
max_row = sheet_kc.max_row
last_empty_row = max_row + 1
for row in range(4, max_row + 1):
    if sheet_kc[f"B{row}"].value is None:
        last_empty_row = row
        break
print(f"⚡ 发现 B 列第一个空单元格所在行: {last_empty_row}")

# 解除合并单元格（排除第一行）
merged_cells_ranges = list(sheet_kc.merged_cells.ranges)
for merged_range in merged_cells_ranges:
    if merged_range.min_row > 1:
        sheet_kc.unmerge_cells(str(merged_range))

# 插入列
sheet_kc.insert_cols(10, 10)
sheet_kc.insert_cols(3, 1)

# C列居中
for cell in sheet_kc["C"]:
    cell.alignment = Alignment(horizontal="center", vertical="center")

# 合并标题
sheet_kc.merge_cells('H3:J3')
sheet_kc.merge_cells('U3:W3')
sheet_kc["U3"] = "不合格"

# 提取第2列数据后5位写入第3列
for row in sheet_kc.iter_rows(min_row=2, max_row=last_empty_row - 1, max_col=3):
    cell_value = str(row[1].value).strip() if row[1].value else ""
    match = re.search(r'\d+', cell_value)
    if match:
        row[2].value = match.group()[-5:].zfill(5)  # ✅ 修改为后5位，补足5位
sheet_kc["C4"] = "编号"


# 设置 K4:O4 表头
header_titles = ["外应存", "最小发货", "家里库存", "家应存", "排产", "月计划", "月计划缺口", "外仓出库总量", "外仓入库总量"]
for i, title in enumerate(header_titles):
    col_letter = chr(ord('K') + i)
    cell = sheet_kc[f"{col_letter}4"]
    cell.value = title
    cell.fill = PatternFill(start_color="C187F7", end_color="C187F7", fill_type="solid")
# ===========================================
# 📄 处理第一页副本：筛选“成品库”数据生成副本
# ===========================================
if "第一页" in merged_wb.sheetnames:
    sheet_first = merged_wb["第一页"]
    header_row = [cell.value for cell in sheet_first[1]]

    # 如果表头中包含“仓库”字段，则提取属于“成品库”的数据行
    if "仓库" in header_row:
        warehouse_col_index = header_row.index("仓库") + 1

        # 只保留仓库列为“成品库”的行
        data_to_copy = [
            row for row in sheet_first.iter_rows(min_row=2, values_only=True)
            if row[warehouse_col_index - 1] == "成品库"
        ]

        # 创建“第一页副本”工作表，复制表头和筛选后的数据
        sheet_copy = merged_wb.create_sheet("第一页副本")
        sheet_copy.append(header_row)
        for row in data_to_copy:
            sheet_copy.append(row)

# ===========================================
# 📄 创建家里库存表：提取编号+存货名称+数量字段
# ===========================================
if "第一页副本" in merged_wb.sheetnames:
    sheet_copy = merged_wb["第一页副本"]
    header_row = [cell.value for cell in sheet_copy[1]]

    # 如果表头中包含“存货名称”，则开始提取
    if "存货名称" in header_row:
        inventory_name_col_index = header_row.index("存货名称") + 1
        extracted_data = []

        for row in sheet_copy.iter_rows(min_row=2, values_only=True):
            # 处理存货名称字段
            item_name = str(row[inventory_name_col_index - 1]).strip() if row[inventory_name_col_index - 1] else ""

            # 提取前5位作为编号（仅当前5字符不全是中文时）
            five_digits = item_name[:5] if not all('\u4e00' <= char <= '\u9fa5' for char in item_name[:5]) else ""

            # 提取主数量字段（转换为数值）
            quantity = row[header_row.index("主数量")] if "主数量" in header_row else None
            if isinstance(quantity, str):
                quantity = float(quantity) if quantity.replace(".", "", 1).isdigit() else None

            # 收集编号、名称、数量
            extracted_data.append([five_digits, item_name, quantity])

        # 创建“家里库存”工作表，并写入标题和数据
        home_stock_sheet = merged_wb.create_sheet("家里库存")
        home_stock_sheet.append(["编号", "存货名称", "数量"])
        for data in extracted_data:
            home_stock_sheet.append(data)

        # 设置“数量”列为千位分隔格式
        for row in home_stock_sheet.iter_rows(min_row=2, max_row=home_stock_sheet.max_row, min_col=3, max_col=3):
            for cell in row:
                if cell.value is not None:
                    cell.number_format = "#,##0.00"


# ========================================
# 🔁 比对“家里库存”编号，将数量写入第13列（M列）
# ========================================
if "家里库存" in merged_wb.sheetnames:
    home_stock_sheet = merged_wb["家里库存"]

    # 先将“库存表”中编号列（第3列）构造成后4位 → 行映射（用于4位匹配）
    inventory_suffix_dict = {
        str(row[2].value)[-4:]: row for row in sheet_kc.iter_rows(min_row=2, max_row=last_empty_row - 1, max_col=13)
        if row[2].value
    }

    # 正常5位编号 → 直接编号匹配映射
    inventory_code_dict = {
        str(row[2].value).zfill(5): row for row in sheet_kc.iter_rows(min_row=2, max_row=last_empty_row - 1, max_col=13)
        if row[2].value
    }

    # 遍历“家里库存”编号行
    for row in home_stock_sheet.iter_rows(min_row=2, values_only=True):
        raw_code = str(row[0]).strip() if row[0] else ""
        quantity = row[2]

        # 类型1️⃣：编号是 4位数字+“-” 的格式
        if re.fullmatch(r"\d{4}-", raw_code):
            key4 = raw_code[:4]
            if key4 in inventory_suffix_dict:
                inventory_row = inventory_suffix_dict[key4]
                inventory_row[12].value = quantity

        # 类型2️⃣：标准 5位编号匹配
        elif re.fullmatch(r"\d{5}", raw_code):
            key5 = raw_code.zfill(5)
            if key5 in inventory_code_dict:
                inventory_row = inventory_code_dict[key5]
                inventory_row[12].value = quantity


# 会计格式：G到Q列
for col in range(7, 22):
    col_letter = get_column_letter(col)
    format_range = f"{col_letter}5:{col_letter}{last_empty_row - 1}"
    for cell in sheet_kc[format_range]:
        for c in cell:
            c.alignment = Alignment(horizontal="right")
            c.number_format = '_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)'

# 设置列宽
col_widths = {
    'B': 4.5, 'C': 5, 'D': 35.88, 'E': 3, 'F': 3.6,
    'G': 8.6, 'H': 8, 'I': 8, 'J': 8,
    'K': 5.88, 'L': 8.1, 'M': 9.8, 'N': 5.88,
    'O': 9.5, 'P': 9.8, 'Q': 10.08, 'R': 9.5, 'S': 9.5
}
for col, width in col_widths.items():
    sheet_kc.column_dimensions[col].width = width + 0.6
print("✅ 固定列宽设置完成")

# ✅ 设置第1行行高为8
sheet_kc.row_dimensions[1].height = 18
print("✅ 第1行行高已设置为 8")


# 冻结前4行
sheet_kc.freeze_panes = "A5"
print("✅ 表头冻结完成（冻结到第4行）")

# 折叠第2行
sheet_kc.row_dimensions[2].outlineLevel = 1
sheet_kc.row_dimensions[2].hidden = True
sheet_kc.sheet_properties.outlinePr.summaryBelow = True
print("✅ 第2行已折叠（默认隐藏，可展开）")

# 设置缩放比例
sheet_kc.sheet_view.zoomScale = 95
print("✅ 工作表缩放比例已设置为 95%")

# 保存
merged_wb.save(latest_file)
merged_wb.close()
print("🎉 全部处理完成！")

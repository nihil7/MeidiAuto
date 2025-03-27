import openpyxl
import os
import re
from openpyxl.styles import Font, Alignment, Border, Side
import sys

# ================================
# 1. 文件路径配置（支持传参）
# ================================
# 默认路径
default_inventory_folder = os.path.join(os.getcwd(), "data")  # GitHub 使用相对路径

# 判断是否传入路径
if len(sys.argv) >= 2:
    inventory_folder = sys.argv[1]
    print(f"✅ 使用传入路径: {inventory_folder}")
else:
    inventory_folder = default_inventory_folder
    print(f"⚠️ 未传入路径，使用默认路径: {inventory_folder}")

# 确保文件夹路径存在
if not os.path.exists(inventory_folder):
    print(f"❌ 文件夹路径不存在: {inventory_folder}")
    exit()


# 获取当前 Python 脚本所在目录
script_dir = os.path.dirname(os.path.abspath(__file__))

# 获取 `data` 目录的路径
data_folder = os.path.join(script_dir, "data")

# 确保 `data` 目录存在
if not os.path.exists(data_folder):
    print(f"❌ 数据文件夹不存在: {data_folder}")
    exit()

# 获取当前脚本所在目录（即 script/ 目录）
script_dir = os.path.dirname(os.path.abspath(__file__))

# 获取 `data` 目录的正确路径
data_folder = os.path.join(script_dir, "data")

# 确保 `data` 目录存在
if not os.path.exists(data_folder):
    print(f"❌ 数据文件夹不存在: {data_folder}")
    exit()

# 设置 '量化需求' 文件路径
demand_file = os.path.join(data_folder, "list.xlsx")

# 确保文件存在
if not os.path.exists(demand_file):
    print(f"❌ 文件不存在: {demand_file}")
    exit()

print(f"✅ 找到文件: {demand_file}")

# ================================
# 2. 查找“总库存”文件
# ================================
inventory_file = None
for file in os.listdir(inventory_folder):
    if file.endswith('.xlsx') and '总库存' in file:
        inventory_file = os.path.join(inventory_folder, file)
        break

if not inventory_file:
    print("❌ 没有找到包含'总库存'的文件！")
    exit()

print(f"✅ 找到库存文件：{inventory_file}")

# ================================
# 3. 打开Excel文件
# ================================
wb_demand = openpyxl.load_workbook(demand_file)
sheet_demand = wb_demand['2503']

wb_inventory = openpyxl.load_workbook(inventory_file)
sheet_inventory = wb_inventory['库存表']

# ================================
# 4. 提取"2503"数据
# ================================
demand_data = {
    str(row[0]).strip(): (row[1], row[2], row[3])
    for row in sheet_demand.iter_rows(min_row=2, max_col=4, values_only=True)
}

# ================================
# 5. 更新“库存表”数据
# ================================
updated_count = 0
start_row = 5  # 数据起始行

for row in sheet_inventory.iter_rows(min_row=start_row, max_col=15):
    inventory_code = str(row[2].value).strip()

    if inventory_code in demand_data:
        B值, C值, D值 = demand_data[inventory_code]

        sheet_inventory.cell(row=row[0].row, column=11, value=B值)  # K列
        sheet_inventory.cell(row=row[0].row, column=14, value=C值)  # N列
        sheet_inventory.cell(row=row[0].row, column=16, value=D值)  # P列

        updated_count += 1

# ================================
# 6. 格式化单元格
# ================================
def set_alignment(sheet, min_row, min_col, max_col, align='right'):
    """设置对齐方式"""
    for col in range(min_col, max_col + 1):
        for row in sheet.iter_rows(min_row=min_row, min_col=col, max_col=col):
            for cell in row:
                cell.alignment = Alignment(horizontal=align)

set_alignment(sheet_inventory, min_row=start_row, min_col=11, max_col=17)

# ================================
# 7. 设置K4:Q49单元格的边框和字体大小
# ================================
thin_border = Border(
    top=Side(style="thin"),
    left=Side(style="thin"),
    right=Side(style="thin"),
    bottom=Side(style="thin")
)

for row in sheet_inventory.iter_rows(min_row=4, max_row=51, min_col=6, max_col=17):
    for cell in row:
        cell.border = thin_border
        cell.font = Font(size=10)

for row in sheet_inventory.iter_rows(min_row=4, max_row=100, min_col=11, max_col=11):
    for cell in row:
        cell.font = Font(size=8)

for row in sheet_inventory.iter_rows(min_row=4, max_row=100, min_col=14, max_col=14):
    for cell in row:
        cell.font = Font(size=8)

# ================================
# 8. 合并单元格R3:T3并写入“不合格”
# ================================
sheet_inventory.merge_cells('R3:T3')
sheet_inventory['R3'] = '不合格'

# ================================
# 9. 检查是否含有汉字，设置字体大小为5
# ================================
def contains_chinese(text):
    """判断字符串是否包含汉字"""
    return bool(re.search('[\u4e00-\u9fff]', str(text)))

def modify_font_size_if_chinese(sheet, col, min_row=5, font_size=5):
    """检查是否有汉字，设置字体大小"""
    for row in sheet.iter_rows(min_row=min_row, min_col=col, max_col=col):
        for cell in row:
            if contains_chinese(cell.value):
                print(f"✅ 发现汉字：{cell.value}，单元格位置：{cell.coordinate}")
                cell.font = Font(size=font_size)

modify_font_size_if_chinese(sheet_inventory, col=11)  # K列
modify_font_size_if_chinese(sheet_inventory, col=14)  # N列

# ================================
# 10. **保存源文件**
# ================================
try:
    wb_inventory.save(inventory_file)
    print(f"✅ 完成！共更新 {updated_count} 行。")
    print(f"✅ 源文件已直接修改并保存：{inventory_file}")
except Exception as e:
    print(f"❌ 保存源文件失败：{e}")

import os
import sys
import glob
import openpyxl
from openpyxl.styles import PatternFill


# ================================
# 📂 文件路径配置（支持主程序传参）
# ================================

default_inventory_folder = os.path.abspath(os.path.join(os.getcwd(), "data"))

# 支持外部传参路径（来自主程序）
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

# ================================
# 1. 文件查找和筛选
# ================================

# 匹配文件：总库存*.xlsx
pattern = os.path.join(inventory_folder, '总库存*.xlsx')
files = glob.glob(pattern)

# 过滤掉 Excel 的临时文件（以~$开头）
valid_files = [f for f in files if not os.path.basename(f).startswith('~$')]

# 判断文件是否找到
if not valid_files:
    print("❌ 没有找到符合条件的文件！")
    sys.exit(1)

# 取第一个有效文件
inventory_file = valid_files[0]
print(f"✅ 找到文件：{inventory_file}")

# ================================
# ✅ 后续可以继续处理 Excel 文件
# ================================
# ================================
# 2. 打开Excel文件，读取工作表
# ================================
wb_inventory = openpyxl.load_workbook(inventory_file)
sheet = wb_inventory['库存表']  # 假设工作表名为 '库存表'

# ================================
# 3. 数据处理与填充颜色
# ================================
from openpyxl.styles import PatternFill

def process_inventory_data(sheet):
    # 主颜色填充定义
    purple_fill = PatternFill(start_color="3F0065", end_color="3F0065", fill_type="solid")  # 深紫
    red_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")     # 红色
    green_fill = PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid")   # 绿色

    # 扩展颜色（用于同行其他单元格）
    purple_row_fill = PatternFill(start_color="CCC0DA", end_color="CCC0DA", fill_type="solid")  # 淡紫
    red_row_fill = PatternFill(start_color="E6B8B7", end_color="E6B8B7", fill_type="solid")      # 淡红

    # 遍历所有行（跳过表头）
    for row in sheet.iter_rows(min_row=2, max_col=12, values_only=False):
        column_10 = row[9]   # 第10列（索引9）
        column_12 = row[11]  # 第12列（索引11）

        # 安全转换数值
        try:
            n = float(column_12.value) if column_12.value is not None else 0
        except ValueError:
            n = 0
        try:
            m = float(column_10.value) if column_10.value is not None else 0
        except ValueError:
            m = 0

        if n > 0:  # 有需求才处理

            # 情况 1：填紫色
            if (m == 0 and n > 0) or m < 0:
                print(f"行 {row[0].row} → 紫色: n={n}, m={m}")
                column_12.fill = purple_fill

                # 给当前行除第12列以外、有值的单元格染淡紫色
                for i, cell in enumerate(row):
                    if i != 11 and cell.value is not None:  # 排除第12列，避免覆盖
                        cell.fill = purple_row_fill

            # 情况 2：填绿色（不扩展同行填色）
            elif m != 0 and n / m < 1:
                print(f"行 {row[0].row} → 绿色: n={n}, m={m}")
                column_12.fill = green_fill

            # 情况 3：填红色 + 扩展染色
            elif m != 0 and n / m >= 1:
                print(f"行 {row[0].row} → 红色: n={n}, m={m}")
                column_12.fill = red_fill

                for i, cell in enumerate(row):
                    if i != 11 and cell.value is not None:
                        cell.fill = red_row_fill



# ================================
# 4. 主程序
# ================================
def main(folder_path):
    pattern = os.path.join(folder_path, '总库存*.xlsx')
    files = glob.glob(pattern)
    if not files:
        print("❌ 没有找到符合条件的文件！")
        return

    inventory_file = files[0]  # 取第一个文件
    wb_inventory = openpyxl.load_workbook(inventory_file)
    sheet = wb_inventory['库存表']

    process_inventory_data(sheet)

    # 保存Excel文件，直接覆盖原文件
    wb_inventory.save(inventory_file)  # 保存到原文件路径
    print(f"✅ 处理后的文件已保存为：{inventory_file}")

# ================================
# 调用主程序
# ================================
main(inventory_folder)

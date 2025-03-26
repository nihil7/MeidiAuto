import os
import sys
import glob
import openpyxl
from openpyxl.styles import Font


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

# ================================
# 1. 查找文件（过滤临时文件）
# ================================

# 搜索文件名匹配的文件（支持通配符）
pattern = os.path.join(inventory_folder, '总库存*.xlsx')
files = glob.glob(pattern)

# 过滤掉临时文件（通常以 ~$ 开头）
valid_files = [f for f in files if not os.path.basename(f).startswith('~$')]

# 确保至少找到一个符合条件的文件
if not valid_files:
    print("❌ 没有找到符合条件的文件！")
    sys.exit(1)

# 取第一个有效文件
inventory_file = valid_files[0]

print(f"✅ 找到文件：{inventory_file}")

# ================================
# 你可以在这里继续执行后续操作
# ================================
# 例如：
# wb_inventory = openpyxl.load_workbook(inventory_file)
# print(wb_inventory.sheetnames)

# ================================
# 2. 打开Excel文件，读取工作表
# ================================
try:
    wb_inventory = openpyxl.load_workbook(inventory_file)
    print(f"✅ 成功加载文件：{inventory_file}")
except Exception as e:
    print(f"❌ 读取 Excel 失败：{e}")
    exit()

sheet_name = "库存表"
if sheet_name not in wb_inventory.sheetnames:
    print(f"❌ 未找到工作表：{sheet_name}")
    exit()

sheet = wb_inventory[sheet_name]
print(f"✅ 成功读取工作表：{sheet_name}")

# ================================
# 3. 读取 "库存表"，查找表头
# ================================
header_row = sheet[4]  # 假设标题在第4行
headers = {}

# 遍历表头行，存储列名和列索引
for cell in header_row:
    if cell.value:
        headers[cell.value.strip()] = cell.column

# 必须存在的列
required_columns = ["外应存", "家应存", "家里库存", "最小发货", "排产"]
missing_columns = [col for col in required_columns if col not in headers]

if missing_columns:
    print(f"❌ 缺少必要的列: {missing_columns}")
    exit()


# 获取列索引（列号转换为 Excel 字母，如 C、D、E）
def col_letter(col_num):
    return openpyxl.utils.get_column_letter(col_num)


col_external = headers["外应存"]
col_home = headers["家应存"]
col_stock = headers["家里库存"]
col_min_ship = headers["最小发货"]
col_production = headers["排产"]
col_ref = 10  # 参照列（第10列）

print("✅ 成功获取表头索引，开始写入公式并处理颜色...")

# ================================
# 4. 遍历数据行并写入 Excel 公式 & 设置颜色
# ================================
gray_font = Font(color="D8D8D8")  # 灰色字体 (216,216,216)
default_font = Font(color="000000")  # 默认黑色字体

for row_idx in range(5, sheet.max_row + 1):  # 从第5行开始遍历
    # 获取 Excel 单元格地址（A1 格式）
    cell_external = f"{col_letter(col_external)}{row_idx}"  # 外应存
    cell_home = f"{col_letter(col_home)}{row_idx}"  # 家应存
    cell_stock = f"{col_letter(col_stock)}{row_idx}"  # 家里库存
    cell_ref = f"{col_letter(col_ref)}{row_idx}"  # 参照列（第10列）

    cell_min_ship = sheet[f"{col_letter(col_min_ship)}{row_idx}"]  # 最小发货单元格
    cell_production = sheet[f"{col_letter(col_production)}{row_idx}"]  # 排产单元格


    # 读取数值，并进行 float 转换，确保是数字
    def safe_float(value):
        try:
            return float(value) if value is not None and value != "" else 0
        except ValueError:  # 如果转换失败，返回 0
            return 0


    external_stock = safe_float(sheet[cell_external].value)  # 外应存
    home_stock = safe_float(sheet[cell_home].value)  # 家应存
    stock_at_home = safe_float(sheet[cell_stock].value)  # 家里库存
    ref_value = safe_float(sheet[cell_ref].value)  # 参照列值

    # 计算最小发货和排产（Python 侧，作为初始值）
    min_ship_result = external_stock - ref_value  # 计算最小发货
    production_result = home_stock + external_stock - ref_value - stock_at_home  # 计算排产

    # 将计算结果写入单元格，避免公式
    cell_min_ship.value = min_ship_result  # 写入最小发货计算结果
    cell_production.value = production_result  # 写入排产计算结果

    # **动态设置字体颜色**
    if min_ship_result <= 0:
        cell_min_ship.font = gray_font  # 设为灰色
    else:
        cell_min_ship.font = default_font  # 设为黑色

    if production_result <= 0:
        cell_production.font = gray_font  # 设为灰色
    else:
        cell_production.font = default_font  # 设为黑色

print("✅ 公式已替换为数值，正在保存文件...")

# ================================
# 5. 设置特定列的列宽
# ================================
sheet.column_dimensions['K'].width = 3.5
sheet.column_dimensions['O'].width = 7.5
sheet.column_dimensions['B'].width = 8
sheet.column_dimensions['C'].width = 6
sheet.column_dimensions['D'].width = 36.88
sheet.column_dimensions['E'].width = 3
sheet.column_dimensions['H'].width = 5.88
sheet.column_dimensions['I'].width = 5
sheet.column_dimensions['J'].width = 6
sheet.column_dimensions['G'].width = 5
sheet.column_dimensions['N'].width = 5
sheet.column_dimensions['M'].width = 6.88
sheet.column_dimensions['Q'].width = 8
sheet.column_dimensions['F'].width = 3.6
print("✅ 列宽已设置")

# ================================
# 6. 保存Excel文件
# ================================
wb_inventory.save(inventory_file)
wb_inventory.close()
print(f"🎉 文件已保存: {inventory_file}")

import os
import sys
import glob
import openpyxl
from openpyxl.styles import Font

# ================================
# 📂 文件路径配置（支持主程序传参）
# ================================
default_inventory_folder = os.path.abspath(os.path.join(os.getcwd(), "data"))
inventory_folder = sys.argv[1] if len(sys.argv) >= 2 else default_inventory_folder
print(f"📂 使用路径: {inventory_folder}")

if not os.path.exists(inventory_folder):
    print(f"❌ 路径不存在: {inventory_folder}")
    sys.exit(1)

# ================================
# 1. 查找文件（过滤临时文件）
# ================================
pattern = os.path.join(inventory_folder, '总库存*.xlsx')
valid_files = [f for f in glob.glob(pattern) if not os.path.basename(f).startswith('~$')]

if not valid_files:
    print("❌ 没有找到符合条件的文件！")
    sys.exit(1)

inventory_file = valid_files[0]
print(f"✅ 发现库存文件: {inventory_file}")

try:
    # ================================
    # 2. 读取Excel文件，获取工作表
    # ================================
    wb_inventory = openpyxl.load_workbook(inventory_file)
    sheet_name = "库存表"
    if sheet_name not in wb_inventory.sheetnames:
        print(f"❌ 未找到工作表: {sheet_name}")
        sys.exit(1)
    sheet = wb_inventory[sheet_name]
    print(f"✅ 成功读取工作表: {sheet_name}")

    # ================================
    # 3. 读取表头并检查必要列
    # ================================
    # ⚠️ 解决重复列名问题，只保留首个出现的列名
    headers = {}
    for cell in sheet[4]:
        if cell.value:
            key = cell.value.strip()
            if key not in headers:
                headers[key] = cell.column

    required_columns = [
        "外应存", "家应存", "家里库存", "库存",
        "外仓出库总量", "最小发货", "排产", "月计划", "月计划缺口"
    ]

    missing_columns = [col for col in required_columns if col not in headers]
    if missing_columns:
        print(f"❌ 缺少必要列: {missing_columns}")
        sys.exit(1)

    def col_letter(col_num):
        return openpyxl.utils.get_column_letter(col_num)

    col_external = headers["外应存"]
    col_home = headers["家应存"]
    col_stock = headers["家里库存"]
    col_total_stock = headers["库存"]
    col_external_ship = headers["外仓出库总量"]
    col_min_ship = headers["最小发货"]
    col_production = headers["排产"]
    col_plan = headers["月计划"]
    col_gap = headers["月计划缺口"]
    col_ref = 10  # 默认参照列 J

    print("✅ 表头索引解析完成")

    # ================================
    # 4. 遍历数据行，计算公式并更新颜色
    # ================================
    gray_font = Font(color="D8D8D8")
    default_font = Font(color="000000")

    def safe_float(value):
        try:
            return float(value) if value else 0
        except ValueError:
            return 0

    DEBUG_PRINT = True
    DEBUG_ROWS = []  # 仅打印特定行，如 [10, 15]；空列表打印全部

    for row_idx in range(5, sheet.max_row + 1):
        external_stock = safe_float(sheet[f"{col_letter(col_external)}{row_idx}"].value)
        home_stock = safe_float(sheet[f"{col_letter(col_home)}{row_idx}"].value)
        stock_at_home = safe_float(sheet[f"{col_letter(col_stock)}{row_idx}"].value)
        total_stock = safe_float(sheet[f"{col_letter(col_total_stock)}{row_idx}"].value)
        external_shipped = safe_float(sheet[f"{col_letter(col_external_ship)}{row_idx}"].value)
        month_plan = safe_float(sheet[f"{col_letter(col_plan)}{row_idx}"].value)
        ref_value = safe_float(sheet[f"{col_letter(col_ref)}{row_idx}"].value)

        # ===== 计算公式 =====
        min_ship_result = external_stock - ref_value
        production_result = home_stock + external_stock - ref_value - stock_at_home
        gap_result = month_plan - stock_at_home - total_stock - external_shipped

        # ===== 调试打印 =====
        if DEBUG_PRINT and (not DEBUG_ROWS or row_idx in DEBUG_ROWS):
            print(f"🔍 行 {row_idx} | 月计划: {month_plan:.1f}, 家里库存: {stock_at_home:.1f}, "
                  f"库存: {total_stock:.1f}, 外仓出库总量: {external_shipped:.1f} → 缺口: {gap_result:.1f}")

        # ===== 写入结果 =====
        cell_min_ship = sheet[f"{col_letter(col_min_ship)}{row_idx}"]
        cell_production = sheet[f"{col_letter(col_production)}{row_idx}"]
        cell_gap = sheet[f"{col_letter(col_gap)}{row_idx}"]

        cell_min_ship.value = min_ship_result
        cell_production.value = production_result
        cell_gap.value = gap_result

        cell_min_ship.font = gray_font if min_ship_result <= 0 else default_font
        cell_production.font = gray_font if production_result <= 0 else default_font
        cell_gap.font = gray_font if gap_result <= 0 else default_font

    print("✅ 公式计算完成，正在保存文件...")

    # ================================
    # 5. 调整列宽
    # ================================
    col_widths = {
        'B': 8, 'C': 6, 'D': 36.88, 'E': 3, 'F': 3.6,
        'G': 8.6, 'H': 8.8, 'I': 8.8, 'J': 8.8,
        'K': 5.88, 'L': 8.1, 'M': 9.8, 'N': 5.88,
        'O': 8, 'P': 9.8, 'Q': 10.08
    }
    for col, width in col_widths.items():
        sheet.column_dimensions[col].width = width + 0.6

    print("✅ 列宽调整完成")

    # ================================
    # 6. 保存Excel文件
    # ================================
    wb_inventory.save(inventory_file)
    wb_inventory.close()
    print(f"🎉 文件已保存: {inventory_file}")

except Exception as e:
    print(f"❌ Excel 处理失败: {e}")
    sys.exit(1)

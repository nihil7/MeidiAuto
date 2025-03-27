import os
import sys
import shutil
from datetime import datetime
from openpyxl import load_workbook
import platform

# ================================
# 📂 路径配置（支持主程序传参）
# ================================
# 设定默认路径（根据操作系统调整路径）
if platform.system() == "Windows":
    default_folder_path = os.path.join(os.getcwd(), "data")  # 本地 Windows 用相对路径
else:
    default_folder_path = os.path.join(os.getcwd(), "data")  # GitHub 使用相对路径

# 获取路径（本地或传参）
if len(sys.argv) >= 2:
    folder_path = os.path.join(sys.argv[1])  # 传入路径
    print(f"✅ 使用传入路径: {folder_path}")
else:
    folder_path = default_folder_path  # 本地默认路径
    print(f"⚠️ 未传入路径，使用默认路径: {folder_path}")

# 确保路径存在
if not os.path.exists(folder_path):
    print(f"❌ 文件夹路径不存在: {folder_path}")
    sys.exit(1)

# === 查找目录下所有Excel文件，找到第一个包含"合肥市"的文件 ===
excel_files = [f for f in os.listdir(folder_path) if f.endswith('.xlsx') and "合肥市" in f]
print(f"✅ 在目录 {folder_path} 中找到 {len(excel_files)} 个 文件名包含 '合肥市' 的 Excel 文件：")
for idx, file in enumerate(excel_files, 1):
    print(f"{idx}. {file}")

# 如果没有找到文件，提示并退出
if len(excel_files) == 0:
    print("⚠️ 没有找到文件名包含'合肥市'的Excel文件！")
    exit()

# === 打开第一个含有"合肥市"的文件 ===
file_with_hefei = os.path.join(folder_path, excel_files[0])
print(f"\n✅ 打开文件：{file_with_hefei}")
hefei_wb = load_workbook(file_with_hefei)

# === 查找文件夹中其他所有的Excel文件（排除第一个文件） ===
other_excel_files = [f for f in os.listdir(folder_path) if f.endswith('.xlsx') and f != excel_files[0]]
print(f"✅ 找到 {len(other_excel_files)} 个需要合并的文件：")
for idx, file in enumerate(other_excel_files, 1):
    print(f"{idx}. {file}")

# 如果没有其他文件，提示并退出
if len(other_excel_files) == 0:
    print("⚠️ 文件夹中没有其他 Excel 文件可以合并！")
    exit()

# === 创建新的合并文件 ===
timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
merged_filename = f"总库存{timestamp}.xlsx"
merged_filepath = os.path.join(folder_path, merged_filename)

# 复制“合肥市”文件到新文件
shutil.copy(file_with_hefei, merged_filepath)
print(f"\n✅ 创建了新的合并文件：{merged_filename}")

# === 打开合并文件 ===
merged_wb = load_workbook(merged_filepath)

# === 遍历所有需要合并的文件并复制工作表 ===
for file in other_excel_files:
    file_path = os.path.join(folder_path, file)
    wb = load_workbook(file_path)

    # 复制每个工作表到合并文件
    for sheet_name in wb.sheetnames:
        sheet_from = wb[sheet_name]
        if sheet_name in merged_wb.sheetnames:
            merged_wb.remove(merged_wb[sheet_name])
            print(f"ℹ️ 已删除旧的工作表：{sheet_name}")

        sheet_to = merged_wb.create_sheet(sheet_name)
        for row in sheet_from.iter_rows():
            row_values = [cell.value for cell in row]
            sheet_to.append(row_values)
        print(f"✅ 已复制工作表 '{sheet_name}'")

# === 保存合并后的文件 ===
merged_wb.save(merged_filepath)
merged_wb.close()
print(f"✅ 合并完成，文件已保存：{merged_filepath}")

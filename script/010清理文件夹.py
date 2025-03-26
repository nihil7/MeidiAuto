import os
import shutil
import sys
from datetime import datetime
import platform

# =======================
# 设定默认路径
# =======================
if platform.system() == "Windows":
    # 本地 Windows 运行
    default_folder = os.path.join(os.getcwd(), "data", "mail")  # Windows 用相对路径
else:
    # GitHub 运行
    default_folder = os.path.join(os.getcwd(), "data", "mail")

# =======================
# 获取路径（本地 or 传参）
# =======================
if len(sys.argv) >= 2:
    source_folder = os.path.join(sys.argv[1], "mail")  # GitHub 传参路径
    print(f"✅ 已接收外部传入路径: {source_folder}")
else:
    source_folder = default_folder  # 本地默认路径
    print(f"⚠️ 未传入路径，使用默认路径: {source_folder}")

# 验证路径是否存在
if not os.path.exists(source_folder):
    print(f"❌ 错误：路径不存在！ {source_folder}")
    sys.exit(1)

# =======================
# 目标文件夹路径（存放重命名后的文件）
# =======================
target_folder = os.path.join(source_folder, "re")
os.makedirs(target_folder, exist_ok=True)  # 确保目标文件夹存在
print(f"📁 目标文件夹已准备: {target_folder}")

# =======================
# 遍历并移动文件
# =======================
file_count = 0
for filename in os.listdir(source_folder):
    file_path = os.path.join(source_folder, filename)

    # 只处理文件，忽略子文件夹
    if os.path.isfile(file_path):
        # 获取文件的修改时间
        mod_time = datetime.fromtimestamp(os.path.getmtime(file_path))
        time_prefix = mod_time.strftime('%Y%m%d_%H%M%S')

        # 生成新文件名
        new_filename = f"{time_prefix}_{filename}"
        new_file_path = os.path.join(target_folder, new_filename)

        # 移动文件
        shutil.move(file_path, new_file_path)

        print(f"✅ 已移动并重命名: {filename} → {new_filename}")
        file_count += 1

# =======================
# 完成提示
# =======================
print(f"\n🎉 处理完成！共 {file_count} 个文件已移动到: {target_folder}")

import sys
import os
import glob
import xlwings as xw
from PIL import ImageGrab
from datetime import datetime

# 1. 文件路径配置
# ================================
default_inventory_folder = r'C:\Users\ishel\Desktop\当日库存情况'

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

# 匹配文件：总库存*.xlsx
pattern = os.path.join(inventory_folder, '总库存*.xlsx')
files = glob.glob(pattern)

# 确保文件存在
if not files:
    print("❌ 没有找到符合条件的文件！")
    exit()

# 获取最新文件
latest_file = max(files, key=os.path.getctime)
print(f"✅ 找到最新的文件：{latest_file}")

# ================================
# 2. 打开Excel文件，复制 A1:Q60 区域，保存为图片
# ================================

# 打开 Excel 文件并设置 Excel 不显示
app = xw.App(visible=False)  # 设置 visible=False，防止弹出 Excel 窗口
wb = app.books.open(latest_file)
ws = wb.sheets[0]  # 默认打开第一个工作表

# 选择区域：A1 到 Q60
range_to_save_as_image = ws.range('A1:Q60')

# 复制区域为图片
range_to_save_as_image.api.CopyPicture(Format=2)  # Format=2表示复制为图片格式

# 从剪贴板抓取图像并保存为文件
img = ImageGrab.grabclipboard()
if img:
    # 获取当前时间，命名图片
    current_time = datetime.now().strftime("%Y%m%d_%H%M%S")
    image_filename = f"美的仓储自动化_{current_time}.png"
    image_filepath = os.path.join(inventory_folder, image_filename)

    # 保存图片
    img.save(image_filepath, 'PNG')
    print(f"✅ 图片已保存：{image_filepath}")
else:
    print("❌ 未能从剪贴板获取图片")

# 关闭 Excel 文件
wb.close()
app.quit()  # 退出 Excel 应用程序

import subprocess
import os

# 获取当前脚本的绝对路径，并定位到 scripts 目录
script_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "scripts")
common_folder = os.path.join(os.getcwd(), "data")

print(f"📁 公共文件路径已设置为: {common_folder}\n")

# 确保 data 目录存在
os.makedirs(common_folder, exist_ok=True)

# 定义要执行的子程序列表
subprograms = [
    "020邮箱下2个表.py",
    "021单纯的合并.py",
    "030家里库存数据整理格式优化千位分隔符自动列宽.py",
    "032外仓库存数据分析.py",
    "033量化需求插格式优化特殊部分字体缩小.py",
    "041运算和灰度显示格式优化精准列宽.py",
    "042比较和彩色显示.py",
    "050区域单元格的图片.py",
    "051发邮件含图片和附件.py"
]

# 依次执行子程序
for script in subprograms:
    script_path = os.path.join(script_dir, script)
    print(f"🚀 正在运行 {script} ...")
    result = subprocess.run(["python", script_path, common_folder], capture_output=True, text=True)
    print(result.stdout)
    if result.stderr:
        print(f"⚠️ {script} 执行出错: {result.stderr}")

print("\n🎉 全部子程序执行完成！")

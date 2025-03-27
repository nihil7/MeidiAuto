import subprocess
import os

# 获取当前脚本的绝对路径，并定位到 scripts 目录
script_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "script")
common_folder = os.path.join(os.getcwd(), "data")

print(f"📁 公共文件路径已设置为: {common_folder}\n")

# 确保 data 目录存在
os.makedirs(common_folder, exist_ok=True)

# 定义要执行的子程序列表
subprograms = [
    "020 Email download.py",
    "021 Merge excel.py",
    "030 Warehousing at home.py",
    "032 Warehousing at out.py",
    "033 list insertion.py",
    "041 operation.py",
    "042 Color display.py",
    "051 Send an email.py"
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

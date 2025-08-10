import os
import subprocess
import sys

# 获取当前主程序所在文件夹路径
script_dir = os.path.dirname(os.path.abspath(__file__))

# 定义要执行的子程序列表
subprograms = [
    #"010 clean.py",
    "020 Email download.py",
    "021 Merge excel.py",
    "030 Warehousing at home.py",
    "032 Warehousing at out.py",
    "033 list insertion.py",
    "041 operation.py",
    "042 Color display.py",
    "050 image local.py",
    "050 mailtxt.py",
    "052 send email.py",
    "010 clean.py"
    #"051 Send an email.py"

]

# 依次执行子程序
for script in subprograms:
    script_path = os.path.join(script_dir, script)
    print(f"🚀 正在运行 {script} ...")

    try:
        result = subprocess.run(["python", script_path],  # 去掉 common_folder
                                capture_output=True, text=True, encoding="utf-8", check=True)
        print(result.stdout)
    except subprocess.CalledProcessError as e:
        print(f"❌ {script} 运行失败，退出程序！\n错误信息:\n{e.stderr}")
        sys.exit(1)  # 立即终止整个程序

print("\n🎉 全部子程序执行完成！")

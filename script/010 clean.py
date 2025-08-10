import os

# 获取当前脚本所在目录
script_dir = os.path.dirname(os.path.abspath(__file__))

# 要删除的文件名中包含的关键字
keywords = ["总库存", "美的仓储自动化", "合肥市","存量查询","output.html"]

# 遍历目录及其子目录中的所有文件
for root, dirs, files in os.walk(script_dir):
    for filename in files:
        file_path = os.path.join(root, filename)

        # 检查文件名是否包含任意关键字
        if any(keyword in filename for keyword in keywords):
            try:
                os.remove(file_path)
                print(f"✅ 删除文件: {file_path}")
            except Exception as e:
                print(f"❌ 删除文件 {file_path} 失败: {e}")

print("\n处理完成！")

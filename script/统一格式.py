import os

# ========== 配置区 ==========
# 子程序所在文件夹路径（请修改为你的实际路径）
folder_path = r"C:\Users\ishel\Desktop\编程总库\MeidiAuto\script"
# 是否递归子文件夹
recursive = False
# ========== 配置区 ==========

insert_code = [
    "import sys",
    "try:",
    "    sys.stdout.reconfigure(encoding='utf-8')",
    "except AttributeError:",
    "    pass",
    ""
]

def process_file(file_path):
    with open(file_path, 'r', encoding='utf-8') as f:
        lines = f.readlines()

    # 判断是否已经存在 reconfigure
    already_has = any("sys.stdout.reconfigure" in line for line in lines)
    if already_has:
        print(f"✅ 已存在，跳过：{os.path.basename(file_path)}")
        return

    # 找到第一个非空、非注释、非 shebang 的插入位置
    insert_idx = 0
    for idx, line in enumerate(lines):
        stripped = line.strip()
        if stripped == '' or stripped.startswith('#') or stripped.startswith('"""') or stripped.startswith("'''"):
            continue
        if stripped.startswith(("import", "from")):
            insert_idx = idx + 1
        else:
            break

    # 插入代码
    new_lines = lines[:insert_idx] + [line + '\n' for line in insert_code] + lines[insert_idx:]

    # 写回
    with open(file_path, 'w', encoding='utf-8') as f:
        f.writelines(new_lines)

    print(f"✨ 已插入：{os.path.basename(file_path)}")

def main():
    py_files = []

    if recursive:
        for root, dirs, files in os.walk(folder_path):
            for file in files:
                if file.endswith('.py'):
                    py_files.append(os.path.join(root, file))
    else:
        for file in os.listdir(folder_path):
            if file.endswith('.py'):
                py_files.append(os.path.join(folder_path, file))

    if not py_files:
        print("⚠️ 未找到任何 .py 文件，请检查路径。")
        return

    for file in py_files:
        process_file(file)

    print("\n🎉 批量插入完成，请使用快捷方式或 PyCharm 测试。")

if __name__ == '__main__':
    main()

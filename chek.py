import os
import subprocess

def get_git_root_directory():
    """获取本地 Git 仓库的根目录"""
    try:
        git_root = subprocess.check_output(["git", "rev-parse", "--show-toplevel"], universal_newlines=True).strip()
        return git_root
    except subprocess.CalledProcessError:
        print("当前目录不是 Git 仓库，请检查是否在正确的 Git 仓库目录下。")
        return None

def get_git_remote_url():
    """获取关联的 GitHub 仓库 URL"""
    try:
        remote_url = subprocess.check_output(["git", "remote", "get-url", "origin"], universal_newlines=True).strip()
        return remote_url
    except subprocess.CalledProcessError:
        print("没有找到与远程 GitHub 仓库的连接。")
        return None

def main():
    # 获取本地 Git 仓库根目录
    git_root = get_git_root_directory()
    if not git_root:
        return

    # 获取 GitHub 远程仓库 URL
    remote_url = get_git_remote_url()
    if not remote_url:
        return

    # 显示本地 Git 根目录与 GitHub 仓库 URL
    print(f"本地 Git 仓库根目录: {git_root}")
    print(f"GitHub 仓库 URL: {remote_url}")

    # 提取 GitHub 仓库的项目名称
    github_repo = remote_url.split("/")[-1].replace(".git", "")
    local_project = os.path.basename(git_root)

    print(f"本地项目目录名称: {local_project}")
    print(f"GitHub 仓库名称: {github_repo}")

    # 比较本地目录和 GitHub 仓库名称
    if local_project == github_repo:
        print("本地项目目录与 GitHub 仓库名称匹配。")
    else:
        print("本地项目目录与 GitHub 仓库名称不匹配，请确认是否为正确的项目。")

if __name__ == "__main__":
    main()

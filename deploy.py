# -*- coding: utf-8 -*-
"""
一键部署脚本 - 更新数据并推送到 GitHub Pages

用法: python deploy.py
流程: 读取Excel -> 生成HTML -> git提交并推送 -> GitHub Pages自动更新

前提:
  1. 本地 Excel 数据已更新
  2. Git 已配置好 SSH 密钥（已绑定 GitHub）
  3. GitHub 仓库已创建且 remote 已设置
"""

import subprocess
import sys
import os

# 复用 process_data.py 的数据处理逻辑
from process_data import main as process_data

OUTPUT_DIR = os.path.dirname(os.path.abspath(__file__))


def run_git(*args):
    """执行 git 命令并打印输出"""
    cmd = ["git"] + list(args)
    print(f"  $ {' '.join(cmd)}")
    result = subprocess.run(cmd, cwd=OUTPUT_DIR, capture_output=True, text=True, encoding="utf-8")
    if result.stdout.strip():
        print(f"    {result.stdout.strip()}")
    if result.returncode != 0:
        print(f"    [错误] {result.stderr.strip()}")
        return False
    return True


def main():
    print("=" * 50)
    print("  开始部署: 数据处理 + 推送 GitHub Pages")
    print("=" * 50)

    # Step 1: 处理数据，生成 HTML
    print("\n[1/3] 处理 Excel 数据，生成 HTML 文件...")
    process_data()

    # Step 2: Git 提交
    print("\n[2/3] 提交到 Git...")
    run_git("add", "-A")

    # 生成提交信息（带时间戳）
    from datetime import datetime
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    commit_msg = f"数据更新 {timestamp}"
    run_git("commit", "-m", commit_msg)

    # Step 3: 推送
    print("\n[3/3] 推送到 GitHub...")
    if run_git("push"):
        print("\n" + "=" * 50)
        print("  部署完成!")
        print("  GitHub Pages 链接刷新后即可看到最新数据")
        print("  (首次推送后等待 1-2 分钟生效)")
        print("=" * 50)
    else:
        print("\n  [警告] 推送失败，请检查 Git remote 配置")
        print("  如果尚未设置 remote，请运行:")
        print('    git remote add origin git@github.com:Faye0366/<仓库名>.git')


if __name__ == "__main__":
    main()

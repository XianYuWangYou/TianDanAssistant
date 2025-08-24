#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
TianDanAssistant Wiki同步脚本

该脚本用于将docs目录中的Markdown文档同步到Gitee Wiki仓库中。
使用方法：
1. 确保已在Gitee上启用Wiki功能
2. 配置好Git用户信息
3. 运行: python sync_wiki.py
"""

import os
import shutil
import subprocess
import sys


def check_git_repo():
    """检查当前目录是否为Git仓库"""
    try:
        subprocess.check_output(['git', 'rev-parse', '--git-dir'], 
                              stderr=subprocess.DEVNULL)
        return True
    except subprocess.CalledProcessError:
        return False
    except FileNotFoundError:
        print("错误: 未找到Git命令，请先安装Git")
        return False


def clone_wiki_repo(project_url):
    """
    克隆Wiki仓库
    
    Args:
        project_url (str): 主项目Git URL
    
    Returns:
        str: Wiki仓库路径
    """
    wiki_url = project_url.replace('.git', '.wiki.git')
    wiki_dir = os.path.basename(wiki_url)
    
    if os.path.exists(wiki_dir):
        print(f"更新现有Wiki仓库: {wiki_dir}")
        subprocess.run(['git', 'pull'], cwd=wiki_dir)
    else:
        print(f"克隆Wiki仓库: {wiki_url}")
        subprocess.run(['git', 'clone', wiki_url])
    
    return wiki_dir


def copy_docs_to_wiki(docs_dir, wiki_dir):
    """
    将docs目录中的文档复制到Wiki目录
    
    Args:
        docs_dir (str): 文档目录路径
        wiki_dir (str): Wiki仓库路径
    """
    if not os.path.exists(docs_dir):
        print(f"文档目录 {docs_dir} 不存在")
        return False
    
    # 需要同步的文档文件
    doc_files = [
        'Home.md',
        'Installation.md', 
        'UserGuide.md',
        'Development.md',
        'FAQ.md'
    ]
    
    print("同步文档到Wiki仓库:")
    for doc_file in doc_files:
        src_path = os.path.join(docs_dir, doc_file)
        dst_path = os.path.join(wiki_dir, doc_file)
        
        if os.path.exists(src_path):
            shutil.copy2(src_path, dst_path)
            print(f"  复制 {doc_file}")
        else:
            print(f"  跳过 {doc_file} (文件不存在)")
    
    return True


def commit_and_push_wiki(wiki_dir, commit_message):
    """
    提交并推送Wiki更改
    
    Args:
        wiki_dir (str): Wiki仓库路径
        commit_message (str): 提交信息
    """
    # 添加所有更改
    subprocess.run(['git', 'add', '.'], cwd=wiki_dir)
    
    # 检查是否有更改需要提交
    result = subprocess.run(['git', 'status', '--porcelain'], 
                          cwd=wiki_dir, 
                          capture_output=True, text=True)
    
    if result.stdout.strip():
        # 有更改需要提交
        subprocess.run(['git', 'commit', '-m', commit_message], cwd=wiki_dir)
        print("推送Wiki更改到远程仓库...")
        subprocess.run(['git', 'push'], cwd=wiki_dir)
        print("Wiki同步完成")
    else:
        print("没有需要提交的Wiki更改")


def main():
    """主函数"""
    print("TianDanAssistant Wiki同步工具")
    print("=" * 40)
    
    # 检查是否在Git仓库中
    if not check_git_repo():
        print("错误: 当前目录不是Git仓库")
        sys.exit(1)
    
    # 获取主项目远程URL
    try:
        remote_url = subprocess.check_output(
            ['git', 'remote', 'get-url', 'origin'], 
            text=True).strip()
        print(f"主项目远程URL: {remote_url}")
    except subprocess.CalledProcessError:
        print("错误: 无法获取主项目远程URL")
        sys.exit(1)
    
    # 克隆或更新Wiki仓库
    wiki_dir = clone_wiki_repo(remote_url)
    
    # 复制文档到Wiki目录
    docs_dir = os.path.join(os.getcwd(), 'docs')
    if copy_docs_to_wiki(docs_dir, wiki_dir):
        # 提交并推送Wiki更改
        commit_message = "Update documentation from main project"
        commit_and_push_wiki(wiki_dir, commit_message)
    else:
        print("文档同步失败")


if __name__ == "__main__":
    main()
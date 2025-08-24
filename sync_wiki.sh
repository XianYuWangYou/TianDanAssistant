#!/bin/bash

# TianDanAssistant Wiki同步Shell脚本
# 该脚本用于在Linux/Mac环境下同步更新主项目和Wiki

echo "TianDanAssistant Wiki同步工具"
echo "================================"

# 检查Git是否可用
if ! command -v git &> /dev/null; then
    echo "错误: 未找到Git命令，请先安装Git"
    exit 1
fi

# 获取主项目远程URL
GIT_URL=$(git remote get-url origin 2>/dev/null)
if [ $? -ne 0 ] || [ -z "$GIT_URL" ]; then
    echo "错误: 无法获取主项目远程URL"
    exit 1
fi

echo "主项目远程URL: $GIT_URL"

# 构造Wiki仓库URL
WIKI_URL="${GIT_URL/.git/.wiki.git}"
echo "Wiki仓库URL: $WIKI_URL"

# 获取仓库名称
REPO_NAME=$(basename "$GIT_URL" .git)
WIKI_DIR="${REPO_NAME}.wiki"

echo "Wiki目录: $WIKI_DIR"

# 克隆或更新Wiki仓库
if [ -d "$WIKI_DIR" ]; then
    echo "更新现有Wiki仓库..."
    cd "$WIKI_DIR"
    git pull
    cd ..
else
    echo "克隆Wiki仓库..."
    git clone "$WIKI_URL" "$WIKI_DIR"
fi

# 复制文档文件
echo "同步文档到Wiki仓库..."
cp -f docs/*.md "$WIKI_DIR"/

# 提交并推送Wiki更改
echo "提交并推送Wiki更改..."
cd "$WIKI_DIR"
git add .
git commit -m "Update documentation from main project"
git push
cd ..

echo "Wiki同步完成!"
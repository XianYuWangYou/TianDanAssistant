@echo off
REM TianDanAssistant Wiki同步批处理脚本
REM 该脚本用于在Windows环境下同步更新主项目和Wiki

echo TianDanAssistant Wiki同步工具
echo ================================

REM 检查Git是否可用
git --version >nul 2>&1
if %errorlevel% neq 0 (
    echo 错误: 未找到Git命令，请先安装Git
    pause
    exit /b 1
)

REM 获取主项目远程URL
for /f "delims=" %%i in ('git remote get-url origin 2^>nul') do set GIT_URL=%%i
if not defined GIT_URL (
    echo 错误: 无法获取主项目远程URL
    pause
    exit /b 1
)

echo 主项目远程URL: %GIT_URL%

REM 构造Wiki仓库URL
set WIKI_URL=%GIT_URL:.git=.wiki.git%
echo Wiki仓库URL: %WIKI_URL%

REM 获取仓库名称
for /f "delims=" %%i in ('echo %GIT_URL%') do (
    set REPO_NAME=%%~ni
    if "%%~xi"==".git" (
        set REPO_NAME=%%~ni
    ) else (
        for /f "delims=" %%j in ('echo %%~ni') do set REPO_NAME=%%j
    )
)

set WIKI_DIR=%REPO_NAME%.wiki

echo Wiki目录: %WIKI_DIR%

REM 克隆或更新Wiki仓库
if exist "%WIKI_DIR%" (
    echo 更新现有Wiki仓库...
    cd "%WIKI_DIR%"
    git pull
    cd ..
) else (
    echo 克隆Wiki仓库...
    git clone %WIKI_URL% "%WIKI_DIR%"
)

REM 复制文档文件
echo 同步文档到Wiki仓库...
copy /Y "docs\*.md" "%WIKI_DIR%\"

REM 提交并推送Wiki更改
echo 提交并推送Wiki更改...
cd "%WIKI_DIR%"
git add .
git commit -m "Update documentation from main project"
git push
cd ..

echo Wiki同步完成!

pause
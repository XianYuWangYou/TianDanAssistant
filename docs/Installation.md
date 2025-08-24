# 安装指南

本文档将指导您如何安装和配置填单助手 (TianDanAssistant)。

## 系统要求

- Python 3.7 或更高版本
- Windows 7/8/10/11 操作系统
- 至少 50MB 可用磁盘空间

## 安装步骤

### 1. 克隆或下载项目

您可以从Gitee仓库克隆项目或下载ZIP包：

```bash
git clone <项目地址>
```

或者直接下载ZIP包并解压到本地目录。

### 2. 安装依赖库

进入项目目录，运行以下命令安装所需依赖：

```bash
pip install -r requirements.txt
```

### 3. 依赖库说明

项目依赖以下第三方库：

- `python-docx==0.8.11` - 用于处理Word文档
- `docx2pdf==0.1.8` - 用于将Word文档转换为PDF
- `PyPDF2==3.0.1` - 用于PDF文件处理
- `openpyxl==3.1.2` - 用于处理Excel文件

### 4. 验证安装

安装完成后，可以通过以下命令验证是否安装成功：

```bash
python -c "import docx, openpyxl; print('依赖库安装成功')"
```

## 常见问题

### 1. 安装依赖时出现错误

如果在安装依赖时出现错误，请尝试以下方法：

- 升级pip: `python -m pip install --upgrade pip`
- 使用国内镜像源安装: `pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple`

### 2. 运行时出现模块导入错误

确保您在正确的Python环境中运行程序，并且所有依赖库都已正确安装。

### 3. PDF转换功能无法使用

PDF转换功能依赖于Microsoft Word，确保您的系统已安装Microsoft Office Word。

## 首次运行

安装完成后，可以通过以下命令运行程序：

```bash
python document_processor.py
```

程序启动后，您将看到图形用户界面，可以开始使用填单助手的各项功能。
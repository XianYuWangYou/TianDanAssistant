# 填单助手 (TianDanAssistant)

![GitHub](https://img.shields.io/github/license/XianYuWengyou/TianDanAssistant)
![Python](https://img.shields.io/badge/python-3.7%2B-blue)

填单助手是一个自动化文档处理工具，可以自动识别Word和Excel模板中的占位符，并根据用户输入的信息批量生成填写完成的文档。

## 功能特点

- **自动识别占位符**：自动扫描Word(.docx)和Excel(.xlsx)模板中的占位符（格式为`{占位符名称}`）
- **批量处理**：一次性处理多个模板文件，提高工作效率
- **多格式支持**：支持Word文档和Excel表格的占位符替换
- **界面友好**：提供图形用户界面，操作简单直观
- **配置保存**：支持保存常用的输入信息，方便下次使用

## 应用场景

该工具特别适用于需要重复填写大量相似文档的场景，例如：

- 银行贷款申请材料填写
- 各类业务申请表处理
- 标准化文档批量生成
- 表格数据自动填充

## 安装依赖

在使用前，请确保已安装所有必要的依赖库：

```bash
pip install -r requirements.txt
```

主要依赖包括：
- python-docx==0.8.11
- docx2pdf==0.1.8
- PyPDF2==3.0.1
- openpyxl==3.1.2

## 使用方法

1. 运行程序：
   ```bash
   python document_processor.py
   ```

2. 在图形界面中选择模板文件（Word或Excel文件）

3. 程序会自动识别所有模板中的占位符并显示在输入界面中

4. 填写相应的信息到对应占位符中

5. 选择输出目录并生成填写完成的文档

## 配置文件说明

- `config.json`：保存用户输入信息和程序设置
- `schemes.json`：定义不同的处理方案，包括模板文件列表和占位符顺序

## 开发者信息

开发者：咸鱼网友 (XianYuWengyou)

## 许可证

本项目采用MIT许可证，详情请参见[LICENSE](LICENSE)文件。
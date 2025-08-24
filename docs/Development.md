# 开发指南

本文档介绍填单助手 (TianDanAssistant) 的项目结构和开发相关信息。

## 项目结构

```
TianDanAssistant/
├── document_processor.py    # 主程序文件
├── requirements.txt         # 依赖库列表
├── config.json             # 用户配置文件
├── schemes.json            # 方案配置文件
├── images/                 # 图片资源目录
│   ├── 捐赠.jpg            # 捐赠二维码
│   └── 软件预览/           # 界面预览图
├── docs/                   # 文档目录
└── README.md               # 项目说明文件
```

## 核心模块

### document_processor.py

这是项目的核心文件，包含以下主要类和功能：

#### DocumentProcessor 类

负责文档处理的核心逻辑：

- 占位符提取：从Word和Excel文件中提取占位符
- 文档生成：根据模板和用户输入生成新文档
- 方案管理：处理方案的保存和加载

#### 主要方法

- `extract_placeholders_from_docx()`: 从Word文档提取占位符
- `extract_placeholders_from_xlsx()`: 从Excel文档提取占位符
- `find_placeholders_in_text()`: 在文本中查找占位符
- `collect_all_placeholders()`: 收集所有模板文件中的占位符
- `replace_placeholders_in_docx()`: 在Word文档中替换占位符
- `replace_placeholders_in_xlsx()`: 在Excel文档中替换占位符
- `generate_documents()`: 生成文档主方法

#### GUI 类

负责图形用户界面：

- 界面布局：创建和管理各个标签页
- 事件处理：响应用户操作
- 数据绑定：将用户输入与文档生成关联

### 配置文件

#### config.json

存储用户配置和输入历史：

```json
{
  "last_output_dir": "输出目录路径",
  "user_inputs": {
    "方案名称": {
      "占位符名称": "用户输入值"
    }
  },
  "placeholder_configs": {
    "占位符名称": {
      "type": "控件类型(entry/combobox/date)",
      "options": ["选项1", "选项2"] // 仅对combobox类型有效
    }
  }
}
```

#### schemes.json

存储方案配置：

```json
{
  "方案名称": {
    "template_files": [
      "模板文件1路径",
      "模板文件2路径"
    ],
    "placeholder_order": [
      "占位符1",
      "占位符2"
    ]
  }
}
```

## 技术细节

### 占位符处理

程序使用正则表达式 `r'\{([^}]+)\}'` 来识别占位符。

对于Word文档，由于docx库的特性，一个占位符可能会被分割到多个run中，程序采用以下策略处理：

1. 将段落中所有run的文本合并
2. 在完整文本中进行占位符查找和替换
3. 清空原有run文本，将替换后的完整文本放入第一个run中

### 多线程处理

对于耗时操作（如PDF转换），程序使用多线程避免界面冻结：

```python
threading.Thread(target=耗时方法, args=(参数元组)).start()
```

### 第三方库

- `python-docx`: 处理Word文档
- `openpyxl`: 处理Excel文档
- `docx2pdf`: Word转PDF
- `PyPDF2`: PDF文件处理
- `tkinter`: 图形界面

## 打包发布

项目使用PyInstaller进行打包：

```bash
pyinstaller --onefile --windowed document_processor.py
```

## 扩展开发

### 添加新的控件类型

1. 在[config.json](file:///d%3A/typecode/TianDanAssistant/config.json)中添加新的控件类型配置
2. 在GUI类中实现对应的控件创建方法
3. 更新保存和加载逻辑以支持新控件类型

### 添加新的文档格式支持

1. 实现对应格式的占位符提取方法
2. 实现对应格式的占位符替换方法
3. 在处理逻辑中添加对新格式的支持

## 代码规范

- 使用中文注释
- 遵循PEP 8代码风格
- 方法和类使用文档字符串说明功能和参数
- 保持代码结构清晰，逻辑分离

## 贡献指南

欢迎提交Issue和Pull Request来改进项目：

1. Fork项目
2. 创建功能分支
3. 提交更改
4. 发起Pull Request
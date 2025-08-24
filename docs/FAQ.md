# 常见问题解答 (FAQ)

本文档收集了使用填单助手 (TianDanAssistant) 过程中可能遇到的常见问题及其解决方案。

## 安装相关问题

### Q: 安装依赖时出现错误怎么办？

A: 可以尝试以下几种方法：

1. 升级pip到最新版本：
   ```bash
   python -m pip install --upgrade pip
   ```

2. 使用国内镜像源安装：
   ```bash
   pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple
   ```

3. 分别安装各个依赖库：
   ```bash
   pip install python-docx
   pip install openpyxl
   pip install docx2pdf
   pip install PyPDF2
   ```

### Q: 运行程序时提示缺少模块怎么办？

A: 确保已正确安装所有依赖库。可以使用以下命令验证：

```bash
python -c "import docx, openpyxl, docx2pdf, PyPDF2; print('所有模块导入成功')"
```

如果仍有问题，请重新安装对应的模块。

## 使用相关问题

### Q: 无法识别模板中的占位符怎么办？

A: 请检查以下几点：

1. 占位符格式是否正确：应使用大括号包围，如`{单位名称}`
2. 占位符是否被分割到多个文本段落中：Word文档中有时一个占位符会被分割到多个run中
3. 模板文件是否损坏：尝试重新保存模板文件
4. 模板格式是否支持：目前支持.docx和.xlsx格式

### Q: 生成的文档中占位符没有被替换怎么办？

A: 请检查以下几点：

1. 是否在录入区填写了对应占位符的信息
2. 占位符名称是否完全匹配（包括大小写和空格）
3. 模板中的占位符是否格式正确

### Q: 程序运行时报错"没有找到指定模块"怎么办？

A: 这通常是由于缺少依赖库导致的。请确保已安装所有依赖库：

```bash
pip install -r requirements.txt
```

### Q: PDF转换功能无法使用怎么办？

A: PDF转换功能依赖Microsoft Word，请检查：

1. 系统是否安装了Microsoft Word
2. Word是否能正常打开
3. 是否在生成文档时勾选了"转换为PDF"选项

## 方案配置问题

### Q: 如何删除不需要的方案？

A: 目前可以通过直接编辑[schemes.json](file:///d%3A/typecode/TianDanAssistant/schemes.json)文件删除不需要的方案。找到对应的方案名称部分，删除整个方案配置块。

### Q: 如何修改已有方案中的模板文件？

A: 有两种方法：

1. 在"方案配置"标签页中加载该方案，删除不需要的模板，添加新模板，然后重新保存方案
2. 直接编辑[schemes.json](file:///d%3A/typecode/TianDanAssistant/schemes.json)文件，修改对应方案的`template_files`数组

### Q: 如何调整占位符的顺序？

A: 在"方案配置"标签页中加载方案后，使用占位符列表旁的"上移"和"下移"按钮调整顺序，然后保存方案。

## 模板制作问题

### Q: 如何在模板中添加新的占位符？

A: 在"模板制作"标签页中：

1. 选择对应的模板文件
2. 在"添加新占位符"输入框中输入新占位符名称
3. 点击"添加占位符"按钮
4. 在打开的文档中将占位符复制到需要的位置
5. 保存文档和模板信息

### Q: 如何删除模板中不需要的占位符？

A: 在"模板制作"标签页中：

1. 选择对应的模板文件
2. 在占位符列表中选中要删除的占位符
3. 点击"删除占位符"按钮
4. 在文档中删除对应的占位符文本
5. 保存文档和模板信息

## 性能问题

### Q: 处理大量模板时程序响应缓慢怎么办？

A: 当处理大量或较大的模板文件时，程序可能会出现响应缓慢的情况。这是正常的，因为程序需要逐一处理每个文件。建议：

1. 分批处理模板文件
2. 关闭其他占用系统资源的程序
3. 确保有足够的系统内存

### Q: 生成文档时程序假死怎么办？

A: 对于包含复杂格式或大量内容的文档，生成过程可能需要一些时间。程序使用多线程处理耗时操作，但某些操作仍可能导致界面暂时无响应。请耐心等待，不要强制关闭程序。

## 兼容性问题

### Q: 是否支持WPS Office？

A: 程序主要针对Microsoft Office设计，与WPS Office的兼容性可能有限。建议使用Microsoft Office以获得最佳体验。

### Q: 是否支持Mac/Linux系统？

A: 目前程序主要在Windows系统上测试，对Mac/Linux系统的支持有限。PDF转换功能依赖Windows COM组件，可能无法在其他系统上正常工作。

### Q: 支持哪些Python版本？

A: 程序支持Python 3.7及以上版本。建议使用Python 3.8-3.10版本以获得最佳兼容性。

## 其他问题

### Q: 如何备份我的配置和方案？

A: 可以备份以下文件：

- [config.json](file:///d%3A/typecode/TianDanAssistant/config.json)：包含用户输入历史和配置信息
- [schemes.json](file:///d%3A/typecode/TianDanAssistant/schemes.json)：包含所有方案配置

定期备份这些文件可以防止数据丢失。

### Q: 如何恢复默认设置？

A: 可以删除或重命名[config.json](file:///d%3A/typecode/TianDanAssistant/config.json)和[schemes.json](file:///d%3A/typecode/TianDanAssistant/schemes.json)文件，程序会在下次启动时创建新的默认配置文件。

### Q: 是否可以自定义界面主题？

A: 目前程序使用默认的tkinter界面主题，暂不支持自定义主题。后续版本可能会添加此功能。

如果您遇到了以上未提及的问题，请在Gitee项目页面提交Issue，我们会尽快回复和解决。
# Extract_Title_Footnote
Extract titles and footnotes from Shell/RTF files
# 标题脚注提取工具

## 项目概述
这是一个用于从Shell/RTF文档中提取标题和脚注的工具。

### 🌟 主要特性
- ✅ 自动提取表、图、列表标题（不含副标题）
- ✅ 自动提取表、图、列表标题的脚注内容
- ✅ 支持多行标题和脚注的自动拆分
- ✅ 上标、下标以及百分号自动替换为unicode
- ✅ 编码版本自动识别和替换（&meddra. &whodrug）
- ✅ 未替换XXX提醒
- ✅ 支持自定义最大footnote列数
- ✅ 支持自定义项目编号（用于输出文件的名称）
- ✅ 支持自定义脚注关键词（替代programming/programmer）
- ✅ 支持处理过程实时取消

## 使用方法

### Shell Processor
1. 双击 ExtractTitleFootnote.exe
2. 切换到"Shell Processor"标签页
3. 填写必填字段：
   - **Shell File Path**: 输入 Shell 文件（注意：必须**Clean**版）的完整路径，或点击"Browse..."按钮选择文件
   - **Max Footnote Columns**: 设置脚注最大列数（默认值：7，范围：1-10）
   - **Project ID**: 输入项目编号，用于输出文件名前缀
4. （可选）自定义脚注关键词：
   - 在"Custom Footnote Keywords"区域可以添加一个或多个关键词
   - 这些关键词将作为脚注提取的终止条件（类似 programming/programmer 的作用）
   - 点击"+ Add Keyword"按钮可以新增关键词输入框
   - 关键词不区分大小写
5. 点击"Confirm"按钮开始处理（当所有必填字段都有效时按钮才启用）
6. 在日志窗口中查看处理进度
7. 如需停止处理，可随时点击"Cancel"按钮
8. 可使用 Ctrl+C 复制日志内容

### RTF Processor
1. 双击 ExtractTitleFootnote.exe
2. 切换到"RTF Processor"标签页
3. 填写必填字段：
   - **LOT File Path**: 输入 LOT 文件（.xlsx 格式）的完整路径，或点击"Browse..."按钮选择文件；**RTF文件应位于同一文件夹内**
   - **Project ID**: 输入项目编号，用于输出文件名前缀
4. 点击"Confirm"按钮开始处理（当所有必填字段都有效时按钮才启用）
5. 在日志窗口中查看处理进度
6. 如需停止处理，可随时点击"Cancel"按钮
7. 可使用 Ctrl+C 复制日志内容

### ⚠️ 注意事项
- **互斥执行**: 当任一处理器正在运行时，另一个处理器的 Confirm 按钮将自动禁用，防止同时运行
- **独立输出**: Shell Processor 生成 Excel 和 Docx 文档，RTF Processor 生成合并的 RTF 文档

## 输出结果

### Shell Processor输出

生成`项目编号_TF_Contents.xlsx`文件、`项目编号_LOT.xlsx`文件和`项目编号_title_footnote_shell.docx` 文件。

`项目编号_TF_Contents.xlsx`包含以下列：
- **序号**: 处理顺序编号
- **标题**: 提取的标题
- **脚注**: 提取的脚注
- **脚注匹配状态**: 提取状态（成功/失败原因）
- **备注**: 特殊说明或警告信息
- **Batch**: 均为1
- **Type**: TLF-Table/TLF-Figure/TLF-Listing，根据标题
- **TLF**: TXX/FXX/LXX，根据标题
- **Prgm Name**: 等于Output Name
- **Output Name**: 以.为分割，如果不足2位自动补0
- **title1, title2, ...**: 拆分后的标题行
- **footnote1, footnote2, ...**: 拆分后的脚注行（受最大列数限制）

`项目编号_LOT.xlsx`包含以下列：
- **文件名称**: 等于Output Name，包含超链接
- **表格名称**: 标题
- **备注**: 空值

`项目编号_title_footnote_shell.docx` 包含标题和脚注


### RTF Processor输出

生成 `项目编号_rtf_title_footnote.rtf` 文件。


## Shell Processor编码版本替换功能

### 功能描述
自动识别并替换脚注中符合特定格式的编码版本信息，根据标题内容选择替换值。

### 匹配规则
- **格式要求**: `编码版本：XXXXX。` (X为任意字符，包括中文、英文、数字等)

### 分类关键词

#### MedDRA分类关键词
标题包含以下任一关键词时，替换为 `&meddra.`：
- 非药物治疗
- 手术  
- 病史
- 不良事件
- AE
- 系统器官分类
- SOC
- PT
- 首选语
- 首选术语

#### WHO Drug分类关键词
标题包含以下任一关键词时，替换为 `&whodrug.`：
- 既往用药
- 合并用药
- 药物治疗 (注意：不包括"非药物治疗")
- ATC
- 按治疗分类/化学分类
- 按治疗分类/化学物质

### 优先级规则
- MedDRA优先级高于WHO Drug
- 当标题同时包含两类关键词时，优先匹配MedDRA
- "非药物治疗"虽然同时属于两类，但按MedDRA处理

### 备注更新
对于执行了编码版本替换的记录，在备注列中自动添加说明：
```
更新编码版本为&meddra./&whodrug.
```

## Shell Processor XXX检查功能

### 功能描述
自动检查替换编码后的脚注是否仍然包含"XXX"（不限大小写，数量≥2），如有则在备注中提醒用户更新。

### 检测规则
- **模式匹配**: `[xX]{2,}` (不区分大小写，至少2个连续的X)

### 匹配示例
✅ 匹配的情况：
- `XX`
- `xxx`
- `XXXX`
- `XxX`
- `xXxX`

❌ 不匹配的情况：
- `X` (只有1个X)
- `XY` (不是连续的X)
- `正常文本`

### 备注更新
检测到XXX时，在备注中添加：
```
注意：footnote存在XX，请更新
```

#### Shell Processor最大footnote列数功能
当脚注拆分后的列数超过指定的最大列数时，程序会自动将超出的列合并到最后一列，并使用`(*ESC*){newline}`作为分隔符。

**示例：**
- 用户指定最大列数为3
- 某脚注拆分后有5行内容
- 结果：前2列正常显示，第3列显示"第3行`(*ESC*){newline}`第4行`(*ESC*){newline}`第5行"

用于控制输出表格的列数，避免列数过多导致超出SAS footnote数量上限。

#### Shell Processor自定义脚注关键词功能
用户可以自定义一个或多个关键词，这些关键词将与默认的programming/programmer关键词一起作为脚注提取的终止条件。

**特点：**
- ✅ 默认保留programming/programmer关键词
- ✅ 支持添加多个自定义关键词
- ✅ 关键词不区分大小写
- ✅ 当段落包含任意一个关键词（默认+自定义）时，脚注提取会在此处终止
- ✅ 只保留关键词之前的内容作为脚注


## 版本信息
- 版本: 1.0.0
- 作者: 董天强
- 时间: 2026年3月

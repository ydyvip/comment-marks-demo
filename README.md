# docx_comment_demo

为 DOCX 文档批量添加批注的 Python 工具，支持段落和单元格的统一索引系统，批注数据通过 JSON 格式输入。

## 🚀 优化特性

- **统一索引系统**：段落和单元格使用连续的索引编号
- **智能文本匹配**：支持跨run的精确文本定位
- **单元格支持**：可以直接在表格单元格内添加批注
- **精简数据结构**：移除冗余属性，支持便利函数计算派生属性
- **向后兼容**：保持原有功能的兼容性

## 项目结构

```
docx_comment_demo/
├── add_comments.py              # 主程序：读取 JSON 批注并写入 DOCX
├── docx_to_json.py              # 优化后的文档内容提取工具
├── scripts/                     # 底层 DOCX 操作工具（unpack/pack/comment）
│   ├── comment.py
│   ├── office/
│   └── templates/
└── README.md
```

## 📋 文档内容提取 (优化版)

### 使用优化后的JSON结构

```bash
# 生成优化后的文档JSON（包含段落和单元格信息）
python docx_to_json.py sample.docx sample_optimized.json

# 查看文档摘要
python docx_to_json.py sample.docx --summary

# 只生成JSON，不显示摘要
python docx_to_json.py sample.docx
```

### 优化后的JSON结构

```json
{
  "metadata": {
    "file_name": "sample.docx",
    "paragraph_count": 980,
    "table_count": 122,
    "cell_count": 3363,
    "total_elements": 4343,
    "total_characters": 50212,
    "total_words": 4028
  },
  "content": [
    {
      "index": 0,
      "element_type": "paragraph",
      "text": "",
      "style": "其他",
      "style_name": "Normal"
    },
    {
      "index": 1,
      "element_type": "paragraph", 
      "text": "签署页",
      "style": "其他",
      "style_name": "Normal"
    },
    {
      "index": 980,
      "element_type": "cell",
      "text": "单元格内容",
      "style": "单元格",
      "table_index": 0,
      "row_index": 0,
      "column_index": 0,
      "table_position": "表格1-行1-列1"
    }
  ]
}
```

### 便利函数

```python
from docx_to_json import get_element_by_index, find_elements_by_text, get_document_stats, calculate_derived_properties

# 获取特定索引的元素
element = get_element_by_index(document_data, 5)

# 根据文本搜索元素
results = find_elements_by_text(document_data, "系统", "paragraph")

# 获取文档统计信息
stats = get_document_stats(document_data)

# 计算派生属性
derived = calculate_derived_properties(element)
```

## 🔍 批注 JSON 格式 (优化版)

批注数据是一个 **对象数组**，支持段落和单元格的统一索引：

| 字段             | 类型    | 说明                                               |
|------------------|---------|--------------------------------------------------|
| `element_index`  | int     | 元素索引（段落或单元格，从 0 开始）                |
| `element_type`   | string  | 元素类型：`"paragraph"` 或 `"cell"`               |
| `match_text`     | string  | 在元素内匹配的文本片段                            |
| `match_occurrence| int     | 同元素内第几次出现，默认 1                        |
| `title`          | string  | 批注标题（作为批注作者显示）                     |
| `content`        | string  | 批注正文内容                                      |

### 示例批注

```json
[
  {
    "element_index": 0,
    "element_type": "paragraph",
    "match_text": "签署页",
    "title": "格式检查",
    "content": "请确认签署页格式是否符合规范要求。"
  },
  {
    "element_index": 980,
    "element_type": "cell",
    "match_text": "项目数据",
    "title": "数据审核",
    "content": "此处数据需要补充完整。"
  },
  {
    "element_index": 16,
    "element_type": "paragraph", 
    "match_text": "文档修改摘要",
    "match_occurrence": 1,
    "title": "标题确认",
    "content": "文档修改摘要标题格式正确。"
  }
]
```

## 🚀 快速开始

### 1. 安装依赖

```bash
pip install python-docx defusedxml
```

### 2. 生成文档结构

**查看完整元素结构（推荐）：**

```bash
# 生成优化后的JSON文档
python docx_to_json.py sample.docx document_structure.json

# 查看元素结构
python add_comments.py sample.docx --list-elements --document-json document_structure.json
```

**查看传统段落结构：**

```bash
# 仅查看段落（向后兼容）
python add_comments.py sample.docx --list-paragraphs
```

### 3. 添加批注

**使用优化后的JSON结构：**

```bash
# 使用优化后的批注JSON
python add_comments.py sample.docx output.docx annotations.json --document-json document_structure.json

# 自定义批注作者
python add_comments.py sample.docx output.docx annotations.json --author "审阅者"
```

**传统方式（向后兼容）：**

```bash
# 传统段落索引方式（向后兼容）
python add_comments.py sample.docx output.docx legacy_annotations.json

# 直接传入JSON字符串
python add_comments.py sample.docx output.docx '[{"element_index":0,"element_type":"paragraph","match_text":"开头","title":"开头","content":"需修改"}]'
```

## 💡 使用建议

### 1. 首次使用流程

```bash
# 1. 查看文档结构
python docx_to_json.py your_document.docx --summary

# 2. 生成详细结构
python docx_to_json.py your_document.docx structure.json

# 3. 查看元素列表
python add_comments.py your_document.docx --list-elements --document-json structure.json

# 4. 创建批注JSON
# 编辑 annotations.json 文件

# 5. 添加批注
python add_comments.py your_document.docx output.docx annotations.json --document-json structure.json
```

### 2. 索引计算说明

- **段落索引**：从 0 开始，对应文档中的段落顺序
- **单元格索引**：从段落数量开始，按表格顺序连续编号
- **统一索引**：所有元素（段落+单元格）使用连续的索引编号

### 3. 派生属性计算

优化后的结构不再存储以下冗余属性，而是通过便利函数计算：

- `has_text` - 文本是否为空
- `is_empty` - 元素是否为空
- `content_type` - 内容类型（内容段落/空段落/内容单元格/空单元格）
- `word_count` - 单词数量

## 🧪 测试验证

运行测试脚本验证优化功能：

```bash
python test_optimized_system.py
```

测试覆盖：
- ✅ 优化后的JSON生成
- ✅ 便利函数功能
- ✅ 向后兼容性
- ✅ 批注格式验证

## 🔄 向后兼容性

- **旧版本批注格式**：仍然支持 `para_index` 字段
- **传统模式**：不使用 `--document-json` 参数时回退到段落处理模式
- **便利函数**：提供派生属性计算功能，保持API兼容性

## 📝 高级功能

### 1. 批注策略优化

```python
# 智能批注定位
annotations = [
    {
        "element_index": index,
        "element_type": "cell",
        "match_text": "数据",
        "title": "数据分析",
        "content": "需要进一步分析"
    }
]
```

### 2. 批注统计与验证

```python
# 批注后验证
def validate_annotations(original_doc, annotated_doc):
    # 实现批注验证逻辑
    pass
```

### 3. 批注导出与管理

```python
# 导出批注信息
def export_comments(annotated_docx):
    # 实现批注导出功能
    pass
```

## 注意事项

- `char_index` 是基于文档**全文**（各段以 `\n` 连接）的字符位置
- 批注作者名取自 `title` 字段；若需统一作者，使用 `--author` 参数
- 输出文件可直接用 Microsoft Word 或 WPS 打开查看批注

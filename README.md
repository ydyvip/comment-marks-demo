# comment_marks_demo

基于 `docx-comments` 库为 DOCX 文档批量添加批注的工具。

## 特性

- **精准锚定**：基于 `docx-comments` 库，自动处理 `commentRangeStart/End` 和 `commentReference` 标记
- **Word Online 兼容**：完整支持 OOXML 评论规范（含 `commentsExtended.xml` 线程和 `commentsIds.xml` 持久化）
- **人员身份管理**：自动维护 `people.xml`，确保批注作者正确显示
- **段落批注**：通过匹配文本在全文段落中搜索定位
- **单元格批注**：通过 `table_index`/`row_index`/`col_index` 精确定位表格单元格并添加批注

## 安装

```bash
pip install docx-comments python-docx
```

## 快速开始

### 1. 列出文档段落

```bash
python add_comments.py sample.docx --list-paragraphs
```

### 2. 列出文档表格

```bash
python add_comments.py sample.docx --list-tables
```

### 3. 段落批注 JSON

```json
[
  {
    "match_text": "签署页",
    "title": "格式检查",
    "content": "请确认签署页格式是否符合规范要求。"
  },
  {
    "match_text": "文档修改摘要",
    "match_occurrence": 1,
    "title": "标题确认",
    "content": "文档修改摘要标题格式正确。"
  }
]
```

### 4. 单元格批注 JSON

```json
[
  {
    "match_text": "协同消息反馈",
    "target_type": "cell",
    "table_index": 0,
    "row_index": 1,
    "col_index": 1,
    "match_occurrence": 1,
    "title": "内容检查",
    "content": "请确认此功能描述是否完整。"
  }
]
```

### 5. 添加批注

```bash
# 段落批注
python add_comments.py sample.docx output.docx annotations.json --author 张三

# 单元格批注
python add_comments.py sample.docx output.docx cell_annotations.json --author 张三
```

## 批注 JSON 格式

### 段落模式（默认）

| 字段 | 类型 | 说明 |
|------|------|------|
| `match_text` | string | 在段落内匹配的文本片段 |
| `match_occurrence` | int | 可选，全文第几次出现，默认 1 |
| `title` | string | 批注标题（显示在批注正文前） |
| `content` | string | 批注正文内容 |
| `target_type` | string | 可选，`"paragraph"`（默认）或 `"cell"` |

### 单元格模式（`target_type: "cell"`）

| 字段 | 类型 | 说明 |
|------|------|------|
| `match_text` | string | 在单元格内匹配的文本片段 |
| `target_type` | string | 必须为 `"cell"` |
| `table_index` | int | 可选，表格索引，从 0 开始（不指定则通配全文档搜索） |
| `row_index` | int | 可选，行索引，从 0 开始（不指定则通配全文档搜索） |
| `col_index` | int | 可选，列索引，从 0 开始（不指定则通配全文档搜索） |
| `match_occurrence` | int | 可选，第几次出现，默认 1 |
| `title` | string | 批注标题 |
| `content` | string | 批注正文内容 |

**索引匹配模式**：指定 `table_index`/`row_index`/`col_index` 精确定位到某个单元格。

**通配匹配模式**：不指定索引，程序自动遍历全文档所有表格的所有单元格，查找包含 `match_text` 的第一个单元格。

## 使用示例

```bash
# 查看段落列表
python add_comments.py document.docx --list-paragraphs

# 查看表格列表
python add_comments.py document.docx --list-tables

# 使用 JSON 文件添加批注
python add_comments.py document.docx output.docx annotations.json --author 张三

# 使用内联 JSON
python add_comments.py document.docx output.docx '[{"match_text":"文本","title":"批注","content":"内容"}]'
```

## 依赖

- Python >= 3.9
- docx-comments >= 0.3.0
- python-docx >= 1.0.0
- lxml >= 4.9.0

## 注意事项

- 批注作者名使用 `--author` 参数；若未指定，默认为 `Reviewer`
- 输出文件可直接用 Microsoft Word 或 WPS 打开查看批注
- 单元格索引从 0 开始，可通过 `--list-tables` 查看表格结构确认索引

# comment_marks_demo

基于 `docx-comments` 库为 DOCX 文档批量添加批注的工具。

## 特性

- **精准锚定**：基于 `docx-comments` 库，自动处理 `commentRangeStart/End` 和 `commentReference` 标记
- **Word Online 兼容**：完整支持 OOXML 评论规范（含 `commentsExtended.xml` 线程和 `commentsIds.xml` 持久化）
- **人员身份管理**：自动维护 `people.xml`，确保批注作者正确显示
- **简单 JSON 格式**：通过段落索引和匹配文本定位批注位置

## 安装

```bash
pip install docx-comments python-docx
```

## 快速开始

### 1. 列出文档段落

```bash
python add_comments.py sample.docx --list-paragraphs
```

### 2. 创建批注 JSON

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

### 3. 添加批注

```bash
python add_comments.py sample.docx output.docx annotations.json --author 张三
```

## 批注 JSON 格式

| 字段 | 类型 | 说明 |
|------|------|------|
| `match_text` | string | 在段落内匹配的文本片段 |
| `match_occurrence` | int | 可选，同段落内第几次出现，默认 1 |
| `title` | string | 批注标题（显示在批注正文前） |
| `content` | string | 批注正文内容 |

## 使用示例

```bash
# 查看段落列表
python add_comments.py document.docx --list-paragraphs

# 使用 JSON 文件
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

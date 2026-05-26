# docx_comment_demo

为 DOCX 文档批量添加批注的 Python 工具，批注数据通过 JSON 格式输入。

## 项目结构

```
docx_comment_demo/
├── add_comments.py          # 主程序：读取 JSON 批注并写入 DOCX
├── create_sample_docx.py    # 辅助脚本：生成演示用示例文档
├── sample_annotations.json  # 示例批注数据（JSON 格式）
├── scripts/                 # 底层 DOCX 操作工具（unpack/pack/comment）
│   ├── comment.py
│   ├── office/
│   └── templates/
└── README.md
```

## 批注 JSON 格式

批注数据是一个 **对象数组**，每个对象包含以下字段：

| 字段          | 类型   | 说明                                     |
|-------------|--------|------------------------------------------|
| `char_index` | int    | 批注锚点在文档全文中的字符索引（从 0 开始）|
| `title`     | string | 批注标题（作为批注作者显示）               |
| `content`   | string | 批注正文内容                              |

### 示例

```json
[
  {
    "char_index": 0,
    "title": "文档开头",
    "content": "请检查标题格式是否符合规范。"
  },
  {
    "char_index": 50,
    "title": "内容审核",
    "content": "此处描述不够清晰，建议补充具体数据或示例。"
  }
]
```

## 快速开始

### 1. 安装依赖

```bash
pip install python-docx defusedxml
```

### 2. 生成示例文档（可选）

```bash
python create_sample_docx.py sample.docx
```

运行后会打印文档全文及字符索引，方便您确定批注的 `char_index`。

### 3. 添加批注

**方式一：JSON 文件**

```bash
python add_comments.py sample.docx output.docx sample_annotations.json
```

**方式二：直接传入 JSON 字符串**

```bash
python add_comments.py sample.docx output.docx '[{"char_index":0,"title":"开头","content":"需修改"}]'
```

**自定义批注作者**

```bash
python add_comments.py sample.docx output.docx annotations.json --author "张三"
```

## 在代码中调用

```python
from add_comments import add_comments_from_json

annotations = [
    {"char_index": 0,   "title": "格式问题", "content": "标题层级不统一，请调整。"},
    {"char_index": 100, "title": "内容建议", "content": "建议此处补充数据来源。"},
    {"char_index": 250, "title": "语法错误", "content": "第三段第二句存在语法问题。"},
]

add_comments_from_json(
    input_docx="my_document.docx",
    output_docx="my_document_reviewed.docx",
    annotations=annotations,
    author="审阅者",
)
```

## 注意事项

- `char_index` 是基于文档**全文**（各段以 `\n` 连接）的字符位置
- 建议先运行 `create_sample_docx.py` 打印字符索引，再填写 JSON
- 批注作者名取自 `title` 字段；若需统一作者，使用 `--author` 参数
- 输出文件可直接用 Microsoft Word 或 WPS 打开查看批注

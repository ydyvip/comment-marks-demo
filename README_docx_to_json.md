# docx_to_json.py 使用说明

## 功能简介

`docx_to_json.py` 是一个Python脚本，用于抽取Microsoft Word文档（.docx）中的所有段落内容，并将其导出为结构化的JSON格式。

## 支持的功能

- ✅ 读取.docx文件的所有段落内容
- ✅ 识别段落样式（标题、正文、列表项等）
- ✅ 统计文档基本信息（段落数、字符数、单词数等）
- ✅ 生成结构化JSON输出
- ✅ 支持中文文档处理

## 安装依赖

```bash
pip install python-docx
```

## 使用方法

### 1. 使用示例

```bash
# 示例1：生成默认JSON文件
python docx_to_json.py sample.docx
# 输出：sample_paragraphs.json

# 示例2：指定输出文件名
python docx_to_json.py sample.docx my_document.json

# 示例3：查看文档摘要
python docx_to_json.py sample.docx --summary
```

## 输出JSON格式

生成的JSON文件包含以下结构：

```json
{
  "metadata": {
    "file_name": "sample.docx",
    "file_path": "完整文件路径",
    "paragraph_count": 8,
    "total_characters": 349,
    "total_words": 28
  },
  "paragraphs": [
    {
      "index": 0,
      "text": "段落内容",
      "style": "其他",
      "has_text": true,
      "text_length": 11,
      "type": "内容段落",
      "word_count": 1
    }
    // ... 更多段落
  ],
  "style_distribution": {
    "其他": 4,
    "正文": 4
  },
  "summary": {
    "non_empty_paragraphs": 8,
    "empty_paragraphs": 0,
    "title_paragraphs": 0,
    "max_paragraph_length": 90
  }
}
```

## 字段说明

### metadata（文档元信息）
- `file_name`: 文件名
- `file_path`: 完整文件路径
- `paragraph_count`: 段落总数
- `total_characters`: 总字符数
- `total_words`: 总单词数

### paragraphs（段落列表）
- `index`: 段落索引（从0开始）
- `text`: 段落内容
- `style`: 段落样式（标题、正文、列表项等）
- `has_text`: 是否有文本内容
- `text_length`: 文本长度（字符数）
- `type`: 段落类型（内容段落、空段落）
- `word_count`: 单词数量
- `alignment`: 对齐方式（如果适用）

### style_distribution（样式分布）
- 各样式段落的数量统计

### summary（摘要统计）
- `non_empty_paragraphs`: 非空段落数量
- `empty_paragraphs`: 空段落数量
- `title_paragraphs`: 标题段落数量
- `max_paragraph_length`: 最长段落长度

## 支持的段落样式

- 标题（Title, Heading1-6）
- 正文（Normal）
- 列表项（List Paragraph）
- 引用（Quote）
- 代码（Code）
- 其他样式

## 错误处理

脚本会处理以下常见错误：
- 文件不存在
- 文件格式错误
- 权限问题
- 编码问题

## 性能说明

- 支持大文档处理
- 内存使用优化
- 快速解析段落内容

## 示例输出

```bash
# 运行摘要模式
$ python docx_to_json.py sample.docx --summary

文档信息：
文件名：sample.docx
段落总数：8
非空段落：8
空段落：0
总字符数：349
总单词数：28

前5个段落：
------------------------------------------------------------
 1. 示例文档：季度工作报告
 2. 本报告总结了本季度的主要工作成果与存在的问题。团队在产品研发、客户服务和市场推广三个方面取得了显著进展，但在成本控制和人员培训方面仍有较大改进空间。
 3. 一、工作成果
 4. 本季度共完成新功能开发 12 项，修复线上 Bug 47 个，用户满意度评分从上季度的 82 分提升至 89 分。客户投诉率下降 15%，续约率维持在 94% ...
 5. 二、存在问题

✅ 成功导出 8 个段落到 sample_paragraphs.json
```

## 注意事项

1. 确保Python版本 >= 3.6
2. 需要安装python-docx库
3. 仅支持.docx格式（Word 2007及以后版本）
4. 中文文档需要确保UTF-8编码支持
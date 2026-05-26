# docx_to_json.py 使用说明

## 功能简介

`docx_to_json.py` 是一个Python脚本，用于抽取Microsoft Word文档（.docx）中的所有段落内容和表格单元格内容，并将其导出为结构化的JSON格式。

## 支持的功能

- ✅ 读取.docx文件的所有段落内容
- ✅ 读取.docx文件的所有表格单元格内容
- ✅ 识别段落样式（标题、正文、列表项等）
- ✅ 识别单元格内容类型（内容单元格、空单元格）
- ✅ 统计文档基本信息（段落数、表格数、单元格数、字符数、单词数等）
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
    "table_count": 3,
    "cell_count": 19,
    "total_elements": 27,
    "total_characters": 192,
    "total_words": 25
  },
  "content": [
    {
      "index": 0,
      "text": "段落内容",
      "style": "标题",
      "type": "段落",
      "has_text": true,
      "text_length": 11,
      "content_type": "内容段落",
      "word_count": 1
    },
    {
      "index": 8,
      "table_index": 0,
      "row_index": 0,
      "column_index": 0,
      "text": "单元格内容",
      "style": "单元格",
      "type": "单元格",
      "has_text": true,
      "text_length": 6,
      "table_position": "表格1-行1-列1",
      "content_type": "内容单元格",
      "word_count": 1
    }
    // ... 更多内容元素
  ],
  "style_distribution": {
    "标题": 1,
    "正文": 4,
    "其他": 3,
    "单元格": 19
  },
  "summary": {
    "paragraphs": {
      "total": 8,
      "non_empty": 8,
      "empty": 0,
      "titles": 1,
      "max_length": 19
    },
    "tables": {
      "total": 3,
      "cells": {
        "total": 19,
        "non_empty": 17,
        "empty": 2
      }
    },
    "content_types": {
      "paragraphs": 8,
      "cells": 19
    }
  }
}
```

## 字段说明

### metadata（文档元信息）
- `file_name`: 文件名
- `file_path`: 完整文件路径
- `paragraph_count`: 段落总数
- `table_count`: 表格总数
- `cell_count`: 单元格总数
- `total_elements`: 所有内容元素总数（段落+单元格）
- `total_characters`: 总字符数
- `total_words`: 总单词数

### content（内容列表）
- **段落元素**:
  - `index`: 段落索引（从0开始）
  - `text`: 段落内容
  - `style`: 段落样式（标题、正文、列表项等）
  - `type`: "段落"（固定值）
  - `has_text`: 是否有文本内容
  - `text_length`: 文本长度（字符数）
  - `content_type`: 段落内容类型（内容段落、空段落）
  - `word_count`: 单词数量
  - `alignment`: 对齐方式（如果适用）

- **单元格元素**:
  - `index`: 单元格唯一索引
  - `table_index`: 表格索引（从0开始）
  - `row_index`: 行索引（从0开始）
  - `column_index`: 列索引（从0开始）
  - `text`: 单元格内容
  - `style`: "单元格"（固定值）
  - `type`: "单元格"（固定值）
  - `has_text`: 是否有文本内容
  - `text_length`: 文本长度（字符数）
  - `table_position`: 表格位置描述（如："表格1-行1-列1"）
  - `content_type`: 单元格内容类型（内容单元格、空单元格）
  - `word_count`: 单词数量

### style_distribution（样式分布）
- 各样式内容元素的数量统计（包括"单元格"样式）

### summary（摘要统计）
- `paragraphs`: 段落统计
  - `total`: 段落总数
  - `non_empty`: 非空段落数量
  - `empty`: 空段落数量
  - `titles`: 标题段落数量
  - `max_length`: 最长段落长度
- `tables`: 表格统计
  - `total`: 表格总数
  - `cells`: 单元格统计
    - `total`: 单元格总数
    - `non_empty`: 非空单元格数量
    - `empty`: 空单元格数量
- `content_types`: 内容类型统计
  - `paragraphs`: 段落数量
  - `cells`: 单元格数量

## 支持的内容样式

### 段落样式
- 标题（Title, Heading1-6）
- 正文（Normal）
- 列表项（List Paragraph）
- 引用（Quote）
- 代码（Code）
- 其他样式

### 单元格样式
- 单元格（固定样式，用于所有表格单元格）

### 内容类型
- 段落：内容段落、空段落
- 单元格：内容单元格、空单元格

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
$ python docx_to_json.py document_with_tables.docx --summary

📄 文档信息：
文件名：document_with_tables.docx
段落总数：8
非空段落：8
空段落：0
表格总数：3
单元格总数：19
非空单元格：17
空单元格：2
总字符数：192
总单词数：25

📝 前5个段落：
------------------------------------------------------------
 1. 包含表格的文档测试
 2. 这是一个包含表格的测试文档。
 3. 下面会有不同类型的表格示例。
 4. 表格1：基本信息
 5. 表格2：项目进度

📊 表格信息：
----------------------------------------
表格1: 3行 × 2列
表格2: 3行 × 3列
表格3: 2行 × 2列

📋 示例表格（表格1）：
----------------------------------------
行1: 列1: '姓名' | 列2: '张三'
行2: 列1: '部门' | 列2: '技术部'
行3: 列1: '职位' | 列2: '工程师'

✅ 成功导出 27 个内容元素到 document_with_tables_content.json
   - 段落: 8 个
   - 单元格: 19 个
   - 表格: 3 个
```

## 注意事项

1. 确保Python版本 >= 3.6
2. 需要安装python-docx库
3. 仅支持.docx格式（Word 2007及以后版本）
4. 中文文档需要确保UTF-8编码支持
5. 表格中的空单元格会被识别为"空单元格"类型
6. 每个表格单元格都会作为一个独立的内容元素进行提取
7. 单元格的位置信息以"表格X-行Y-列Z"格式提供
8. 对于大型文档，提取所有内容可能需要较多内存
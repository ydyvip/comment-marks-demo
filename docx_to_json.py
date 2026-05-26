"""
docx_to_json.py - 抽取DOCX文档的段落内容并导出为JSON文件

功能：
- 读取DOCX文档的所有段落内容
- 导出为JSON格式，包含段落索引和内容
- 支持段落样式信息（标题、正文等）

使用方法：
python docx_to_json.py input.docx output.json
"""

import json
import sys
from pathlib import Path
from docx import Document


def extract_paragraphs_to_json(docx_path: str, output_path: str = None):
    """
    从DOCX文档中提取所有段落内容并保存为JSON
    
    Args:
        docx_path: 输入的DOCX文件路径
        output_path: 输出的JSON文件路径（可选）
    
    Returns:
        dict: 包含文档信息的字典
    """
    input_path = Path(docx_path)
    if not input_path.exists():
        raise FileNotFoundError(f"找不到文件: {docx_path}")
    
    # 读取DOCX文档
    doc = Document(docx_path)
    
    # 提取段落信息
    paragraphs_data = []
    style_mapping = {
        'Title': '标题',
        'Heading1': '标题1', '1': '标题1',
        'Heading2': '标题2', '2': '标题2',
        'Heading3': '标题3', '3': '标题3',
        'Heading4': '标题4', '4': '标题4',
        'Heading5': '标题5', '5': '标题5',
        'Heading6': '标题6', '6': '标题6',
        'Normal': '正文',
        'List Paragraph': '列表项',
        'Quote': '引用',
        'Code': '代码',
    }
    
    for i, paragraph in enumerate(doc.paragraphs):
        paragraph_info = {
            "index": i,
            "text": paragraph.text.strip() if paragraph.text else "",
            "style": style_mapping.get(paragraph.style.name, '其他'),
            "has_text": bool(paragraph.text),
            "text_length": len(paragraph.text) if paragraph.text else 0
        }
        
        # 添加更多段落属性
        if paragraph.alignment:
            paragraph_info["alignment"] = str(paragraph.alignment)
        
        # 检查是否为空段落
        if not paragraph.text.strip():
            paragraph_info["is_empty"] = True
            paragraph_info["type"] = "空段落"
        else:
            paragraph_info["type"] = "内容段落"
            paragraph_info["word_count"] = len(paragraph.text.split())
        
        paragraphs_data.append(paragraph_info)
    
    # 构建完整的文档信息
    document_info = {
        "metadata": {
            "file_name": input_path.name,
            "file_path": str(input_path.absolute()),
            "paragraph_count": len(paragraphs_data),
            "total_characters": sum(p["text_length"] for p in paragraphs_data),
            "total_words": sum(p["word_count"] for p in paragraphs_data if p["has_text"])
        },
        "paragraphs": paragraphs_data,
        "style_distribution": _get_style_distribution(paragraphs_data),
        "summary": {
            "non_empty_paragraphs": len([p for p in paragraphs_data if p["has_text"]]),
            "empty_paragraphs": len([p for p in paragraphs_data if not p["has_text"]]),
            "title_paragraphs": len([p for p in paragraphs_data if p["style"].startswith('标题')]),
            "max_paragraph_length": max([p["text_length"] for p in paragraphs_data if p["has_text"]], default=0)
        }
    }
    
    # 输出到JSON文件
    if output_path:
        output_file = Path(output_path)
        with open(output_file, 'w', encoding='utf-8') as f:
            json.dump(document_info, f, ensure_ascii=False, indent=2)
        print(f"✅ JSON文件已生成: {output_file}")
    
    return document_info


def _get_style_distribution(paragraphs_data):
    """统计段落样式分布"""
    style_count = {}
    for para in paragraphs_data:
        style = para["style"]
        style_count[style] = style_count.get(style, 0) + 1
    return style_count


def print_document_summary(docx_path: str):
    """打印文档摘要信息"""
    try:
        doc = Document(docx_path)
        print(f"\n文档信息：")
        print(f"文件名：{Path(docx_path).name}")
        print(f"段落总数：{len(doc.paragraphs)}")
        
        # 统计信息
        non_empty_count = 0
        total_chars = 0
        total_words = 0
        
        for para in doc.paragraphs:
            if para.text.strip():
                non_empty_count += 1
                total_chars += len(para.text)
                total_words += len(para.text.split())
        
        print(f"非空段落：{non_empty_count}")
        print(f"空段落：{len(doc.paragraphs) - non_empty_count}")
        print(f"总字符数：{total_chars}")
        print(f"总单词数：{total_words}")
        
        # 显示前5个段落
        print(f"\n前5个段落：")
        print("-" * 60)
        for i, para in enumerate(doc.paragraphs[:5]):
            preview = para.text[:80].replace("\n", " ") + "..." if len(para.text) > 80 else para.text
            print(f"{i+1:2d}. {preview}")
            
    except Exception as e:
        print(f"❌ 读取文档失败: {e}")


def main():
    """主函数"""
    if len(sys.argv) < 2:
        print("使用方法：")
        print("  python docx_to_json.py input.docx output.json")
        print("  python docx_to_json.py input.docx --summary")
        print("  python docx_to_json.py input.docx")
        print("\n选项：")
        print("  input.docx  - 输入的DOCX文件路径")
        print("  output.json - 输出的JSON文件路径（可选）")
        print("  --summary   - 只显示文档摘要信息")
        return
    
    docx_path = sys.argv[1]
    
    # 检查文件是否存在
    if not Path(docx_path).exists():
        print(f"❌ 文件不存在: {docx_path}")
        return
    
    # 处理不同命令行参数
    if len(sys.argv) == 3:
        if sys.argv[2] == "--summary":
            print_document_summary(docx_path)
        else:
            output_path = sys.argv[2]
            try:
                result = extract_paragraphs_to_json(docx_path, output_path)
                print(f"\n✅ 成功导出 {len(result['paragraphs'])} 个段落到 {output_path}")
            except Exception as e:
                print(f"❌ 导出失败: {e}")
    elif len(sys.argv) == 2:
        # 没有输出路径，生成默认的JSON文件名
        input_path = Path(docx_path)
        default_output = f"{input_path.stem}_paragraphs.json"
        try:
            result = extract_paragraphs_to_json(docx_path, default_output)
            print(f"\n✅ 成功导出 {len(result['paragraphs'])} 个段落到 {default_output}")
        except Exception as e:
            print(f"❌ 导出失败: {e}")
    else:
        print("❌ 参数错误。使用 --summary 获取帮助信息")


if __name__ == "__main__":
    main()
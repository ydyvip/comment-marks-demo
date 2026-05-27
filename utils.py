"""
utils.py - DOCX 批注系统的公共工具函数

功能：
- 统一的文本提取和处理方式
- 统一的段落和单元格索引计算
- 统一的文本标准化逻辑
- 公共的文档操作函数

使用方法：
from utils import extract_paragraph_text, normalize_text, calculate_element_index, ...
"""

import re
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import List, Dict, Optional, Tuple

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def extract_paragraph_text(para) -> str:
    """
    从段落 XML 元素中提取文本内容（与 add_comments.py 保持一致）
    
    Args:
        para: 段落 XML 元素
    
    Returns:
        str: 段落的完整文本内容
    """
    text_parts = []
    for run in para.findall(f".//{{{W}}}r"):
        t = run.find(f"{{{W}}}t")
        if t is not None and t.text:
            text_parts.append(t.text)
        else:
            # 处理空 run 的情况
            text_parts.append("")
    return "".join(text_parts)


def normalize_text(text: str) -> str:
    """
    标准化文本，用于匹配比较（与 add_comments.py 保持一致）
    
    Args:
        text: 原始文本
    
    Returns:
        str: 标准化后的文本
    """
    if not text:
        return ""
    
    # 移除多余空格，换行符转为空格
    normalized = re.sub(r'\s+', ' ', text.strip())
    normalized = normalized.replace('\n', ' ')
    
    # 保持原始大小写，不做大小写转换
    return normalized


def extract_paragraph_style(para) -> str:
    """
    提取段落样式名称
    
    Args:
        para: 段落 XML 元素
    
    Returns:
        str: 样式名称
    """
    pPr = para.find(f"{{{W}}}pPr")
    if pPr is None:
        return ""
    pStyle = pPr.find(f"{{{W}}}pStyle")
    return pStyle.get(f"{{{W}}}val", "") if pStyle is not None else ""


def parse_paragraphs_from_document_xml(doc_xml_path: Path):
    """
    解析 document.xml 并返回段落列表（与 add_comments.py 保持一致）
    
    Args:
        doc_xml_path: document.xml 文件路径
    
    Returns:
        tuple: (tree, root, paragraphs)
    """
    tree = ET.parse(doc_xml_path)
    root = tree.getroot()
    body = root.find(f".//{{{W}}}body")
    paragraphs = body.findall(f".//{{{W}}}p") if body is not None else []
    return tree, root, paragraphs


def calculate_element_index(paragraph_count: int, table_index: int, row_index: int, column_index: int) -> int:
    """
    计算元素的统一索引（段落 + 单元格）
    
    Args:
        paragraph_count: 段落总数
        table_index: 表格索引（从 0 开始）
        row_index: 行索引（从 0 开始）
        column_index: 列索引（从 0 开始）
    
    Returns:
        int: 元素的统一索引
    """
    # 单元格索引从段落数量开始，按表格顺序连续编号
    # 简单的线性编号，避免复杂计算导致索引重复
    cell_index = 0
    current_table_index = 0
    
    # 这里需要根据实际的文档结构来计算
    # 暂时使用简单的递增方式
    return paragraph_count + cell_index


def find_element_by_index(paragraphs: List, doc_data: Dict, element_index: int, element_type: str):
    """
    根据索引和类型查找元素
    
    Args:
        paragraphs: 段落 XML 元素列表
        doc_data: 文档数据（从 docx_to_json.py 生成）
        element_index: 元素索引
        element_type: 元素类型（"paragraph" 或 "cell"）
    
    Returns:
        tuple: (target_paragraph, target_element_info) 或 (None, None)
    """
    if element_type == "paragraph":
        # 查找段落
        if element_index < len(paragraphs):
            return paragraphs[element_index], {
                "index": element_index,
                "element_type": "paragraph",
                "style": extract_paragraph_style(paragraphs[element_index]),
                "text": extract_paragraph_text(paragraphs[element_index])
            }
        else:
            return None, None
    
    elif element_type == "cell":
        # 查找单元格（需要从 doc_data 中获取信息）
        if doc_data:
            for element in doc_data.get("content", []):
                if (element.get("index") == element_index and 
                    element.get("element_type") == "cell"):
                    # 返回对应的段落（单元格在 XML 中仍然是段落）
                    # 由于单元格索引的计算方式，这里需要特殊处理
                    # 实际上单元格数据需要从表格中提取
                    return None, element
        
        return None, None
    
    else:
        return None, None


def locate_text_in_paragraph(para, match_text: str, occurrence: int = 1) -> Tuple[Optional[ET.Element], bool, str]:
    """
    在段落内定位文本（改进版，与 add_comments.py 保持一致）
    
    Args:
        para: 段落 XML 元素
        match_text: 要匹配的文本
        occurrence: 第几次出现（默认 1）
    
    Returns:
        tuple: (target_run, found, matched_text)
    """
    # 获取段落所有 run
    runs = para.findall(f".//{{{W}}}r")
    if not runs:
        return None, False, None
    
    # 构建 run 文本映射
    run_texts = []
    for run in runs:
        t = run.find(f"{{{W}}}t")
        if t is not None and t.text:
            run_texts.append(t.text.strip())
        else:
            run_texts.append("")
    
    # 拼接完整段落文本
    full_text = "".join(run_texts)
    
    # 如果匹配文本为空，返回第一个非空 run
    if not match_text.strip():
        for run in runs:
            t = run.find(f"{{{W}}}t")
            if t is not None and t.text and t.text.strip():
                return run, True, t.text.strip()
        return None, False, None
    
    # 标准化文本
    normalized_match = normalize_text(match_text)
    normalized_full = normalize_text(full_text)
    
    # 查找匹配位置
    match_positions = []
    search_from = 0
    
    while True:
        pos = normalized_full.find(normalized_match, search_from)
        if pos == -1:
            break
        match_positions.append(pos)
        search_from = pos + 1
    
    # 检查是否找到指定次数的匹配
    if len(match_positions) < occurrence:
        return None, False, None
    
    target_pos = match_positions[occurrence - 1]
    
    # 找到覆盖目标位置的 run
    current_pos = 0
    for i, run_text in enumerate(run_texts):
        if not run_text:
            continue
        
        run_start = current_pos
        run_end = current_pos + len(run_text)
        
        # 检查目标位置是否在这个 run 中
        if run_start <= target_pos < run_end:
            return runs[i], True, normalized_match
        
        current_pos = run_end
    
    # 如果跨多个run，返回第一个包含匹配文本开头的run
    for i, run_text in enumerate(run_texts):
        if not run_text:
            continue
        
        run_start = current_pos
        run_end = current_pos + len(run_text)
        
        # 如果匹配文本跨多个run，返回第一个run
        if run_start <= target_pos < run_end:
            return runs[i], True, normalized_match
        
        current_pos = run_end
    
    # 最后的备选方案：找到包含匹配文本的最长run
    best_run = None
    best_length = 0
    
    for i, run in enumerate(runs):
        t = run.find(f"{{{W}}}t")
        if t is not None and t.text:
            run_text = t.text.strip()
            if normalized_match in run_text and len(run_text) > best_length:
                best_run = run
                best_length = len(run_text)
    
    if best_run:
        return best_run, True, normalized_match
    
    return None, False, None


def validate_annotation_target(doc_data: Dict, element_index: int, element_type: str, match_text: str) -> Dict:
    """
    验证批注目标的可用性
    
    Args:
        doc_data: 文档数据
        element_index: 元素索引
        element_type: 元素类型
        match_text: 匹配文本
    
    Returns:
        dict: 验证结果
    """
    if not doc_data:
        return {"valid": False, "error": "文档数据未加载"}
    
    # 查找目标元素
    target_element = None
    for element in doc_data.get("content", []):
        if (element.get("index") == element_index and 
            element.get("element_type") == element_type):
            target_element = element
            break
    
    if not target_element:
        return {
            "valid": False, 
            "error": f"找不到索引={element_index}, 类型={element_type} 的元素"
        }
    
    # 检查匹配文本
    element_text = target_element.get("text", "")
    normalized_match = normalize_text(match_text)
    normalized_element = normalize_text(element_text)
    
    if normalized_match not in normalized_element:
        return {
            "valid": False, 
            "error": f"在目标元素中找不到匹配文本 '{match_text}'",
            "element_text": element_text[:100] + "..." if len(element_text) > 100 else element_text
        }
    
    return {
        "valid": True,
        "element": target_element,
        "match_count": normalized_element.count(normalized_match)
    }


def get_document_style_mapping():
    """
    获取文档样式映射
    
    Returns:
        dict: 样式映射字典
    """
    return {
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


def calculate_derived_properties(element):
    """
    计算元素的派生属性（向后兼容）
    
    Args:
        element: 内容元素（段落或单元格）
    
    Returns:
        dict: 包含派生属性的字典
    """
    derived = {}
    
    # 基本属性
    text = element.get("text", "")
    has_text = bool(text.strip())
    text_length = len(text)
    
    # 派生属性
    derived["has_text"] = has_text
    derived["text_length"] = text_length
    derived["is_empty"] = not has_text
    
    if has_text:
        derived["content_type"] = "内容段落" if element.get("element_type") == "paragraph" else "内容单元格"
        derived["word_count"] = len(text.split())
    else:
        derived["content_type"] = "空段落" if element.get("element_type") == "paragraph" else "空单元格"
        derived["word_count"] = 0
    
    return derived


# ─────────────────────────────────────────────────────────────────
# 便利函数：向后兼容
# ─────────────────────────────────────────────────────────────────

def get_element_by_index(document_data, index):
    """
    根据索引获取元素（向后兼容）
    
    Args:
        document_data: 完整的文档数据
        index: 元素索引
    
    Returns:
        dict: 元素数据或None
    """
    for element in document_data.get("content", []):
        if element.get("index") == index:
            return element
    return None


def find_elements_by_text(document_data, search_text, element_type=None):
    """
    根据文本搜索元素（向后兼容）
    
    Args:
        document_data: 完整的文档数据
        search_text: 搜索文本
        element_type: 元素类型（"paragraph" 或 "cell"），可选
    
    Returns:
        list: 匹配的元素列表
    """
    results = []
    for element in document_data.get("content", []):
        if element_type and element.get("element_type") != element_type:
            continue
        
        if search_text in element.get("text", ""):
            results.append(element)
    
    return results


def get_document_stats(document_data):
    """
    获取文档统计信息（向后兼容）
    
    Args:
        document_data: 完整的文档数据
    
    Returns:
        dict: 统计信息
    """
    content = document_data.get("content", [])
    
    stats = {
        "total_elements": len(content),
        "paragraphs": len([e for e in content if e.get("element_type") == "paragraph"]),
        "cells": len([e for e in content if e.get("element_type") == "cell"]),
        "total_characters": sum(len(e.get("text", "")) for e in content),
        "total_words": sum(len(e.get("text", "").split()) for e in content if e.get("text")),
        "non_empty_elements": len([e for e in content if e.get("text")]),
        "empty_elements": len([e for e in content if not e.get("text")])
    }
    
    return stats


# ─────────────────────────────────────────────────────────────────
# 新增：表格和单元格处理函数
# ─────────────────────────────────────────────────────────────────

def extract_table_cells(doc_xml_path: Path):
    """
    从document.xml中提取所有表格单元格信息
    
    Args:
        doc_xml_path: document.xml文件路径
    
    Returns:
        list: 单元格信息列表，每个元素包含table_index, row_index, col_index, text
    """
    tree = ET.parse(doc_xml_path)
    root = tree.getroot()
    body = root.find(f".//{{{W}}}body")
    if body is None:
        return []
    
    cells = []
    table_index = 0
    
    # 查找所有表格
    tables = body.findall(f".//{{{W}}}tbl")
    for table in tables:
        row_index = 0
        # 查找所有行
        rows = table.findall(f".//{{{W}}}tr")
        for row in rows:
            col_index = 0
            # 查找所有单元格
            cells_in_row = row.findall(f".//{{{W}}}tc")
            for cell in cells_in_row:
                # 提取单元格文本
                cell_text = ""
                for run in cell.findall(f".//{{{W}}}r"):
                    t = run.find(f"{{{W}}}t")
                    if t is not None and t.text:
                        cell_text += t.text
                cells.append({
                    "table_index": table_index,
                    "row_index": row_index,
                    "column_index": col_index,
                    "text": cell_text.strip(),
                    "table_position": f"表格{table_index+1}-行{row_index+1}-列{col_index+1}"
                })
                col_index += 1
            row_index += 1
        table_index += 1
    
    return cells


def find_paragraph_for_cell(paragraphs: List, doc_data: Dict, element_index: int):
    """
    根据单元格的统一索引找到对应的XML段落
    
    Args:
        paragraphs: XML段落列表
        doc_data: 文档数据（包含元数据）
        element_index: 单元格的统一索引
    
    Returns:
        tuple: (target_paragraph, cell_info) 或 (None, None)
    """
    # 如果没有文档数据，无法查找单元格
    if not doc_data:
        return None, None
    
    # 查找该单元格在文档数据中的信息
    target_cell = None
    for element in doc_data.get("content", []):
        if (element.get("index") == element_index and 
            element.get("element_type") == "cell"):
            target_cell = element
            break
    
    if not target_cell:
        return None, None
    
    # 获取段落总数
    paragraph_count = doc_data.get("metadata", {}).get("paragraph_count", 0)
    
    # 单元格索引应该 >= 段落数，如果是这样，说明是直接引用
    if element_index < paragraph_count:
        # 这种情况不应该发生，单元格索引应该 >= 段落数
        return None, None
    
    # 由于单元格在XML中也是以段落形式存在，我们需要遍历段落找到表格内的单元格
    # 这是一个更精确的查找方法
    
    # 首先尝试从doc_data中获取精确的位置信息
    table_index = target_cell.get("table_index", 0)
    row_index = target_cell.get("row_index", 0)
    column_index = target_cell.get("column_index", 0)
    
    # 需要解析XML来找到实际的表格单元格段落
    # 这里我们使用一个更健壮的方法：
    # 1. 重新解析document.xml以获取表格信息
    # 2. 根据表格位置信息找到对应的段落
    
    # 暂时使用一个更保守的方法：如果单元格索引超出段落范围，无法直接映射
    # 在实际应用中，需要更复杂的逻辑来跟踪段落和表格的对应关系
    # 对于现在的情况，我们返回None，让调用者处理这个情况
    
    return None, None


def resolve_element_to_paragraph(paragraphs: List, doc_data: Dict, element_index: int, element_type: str):
    """
    将元素索引解析为对应的XML段落
    
    Args:
        paragraphs: XML段落列表
        doc_data: 文档数据
        element_index: 元素索引
        element_type: 元素类型
    
    Returns:
        tuple: (target_paragraph, element_info) 或 (None, None)
    """
    if element_type == "paragraph":
        # 直接查找段落
        if element_index < len(paragraphs):
            return paragraphs[element_index], {
                "index": element_index,
                "element_type": "paragraph",
                "style": extract_paragraph_style(paragraphs[element_index]),
                "text": extract_paragraph_text(paragraphs[element_index])
            }
        else:
            return None, None
    
    elif element_type == "cell":
        # 查找单元格
        return find_paragraph_for_cell(paragraphs, doc_data, element_index)
    
    else:
        return None, None


def create_element_mapping(paragraphs: List, doc_data: Dict):
    """
    创建元素索引到XML段落的映射关系
    
    Args:
        paragraphs: XML段落列表
        doc_data: 文档数据
    
    Returns:
        dict: 映射关系字典
    """
    mapping = {}
    
    if not doc_data:
        return mapping
    
    # 首先添加段落的映射
    for i, para in enumerate(paragraphs):
        mapping[i] = {
            "paragraph": para,
            "element_type": "paragraph",
            "style": extract_paragraph_style(para),
            "text": extract_paragraph_text(para)
        }
    
    # 然后处理单元格的映射
    content = doc_data.get("content", [])
    paragraph_count = doc_data.get("metadata", {}).get("paragraph_count", 0)
    
    # 找出所有单元格元素
    cell_elements = [e for e in content if e.get("element_type") == "cell"]
    
    # 对于单元格，我们需要一个更智能的映射策略
    # 由于单元格在XML中也是以段落形式存在，我们可以尝试根据内容匹配
    for cell in cell_elements:
        element_index = cell.get("index")
        cell_text = cell.get("text", "")
        
        # 如果单元格文本为空，跳过映射
        if not cell_text:
            continue
        
        # 在XML段落中查找匹配的单元格内容
        for i, para in enumerate(paragraphs):
            # 跳过已经映射的段落
            if i in mapping:
                continue
            
            para_text = extract_paragraph_text(para)
            if para_text == cell_text:
                # 找到匹配的段落
                mapping[element_index] = {
                    "paragraph": para,
                    "element_type": "cell",
                    "style": "单元格",
                    "text": cell_text,
                    "table_position": cell.get("table_position", "N/A"),
                    "original_index": element_index
                }
                break
    
    return mapping


def find_element_by_mapping(mapping: Dict, element_index: int, element_type: str):
    """
    使用映射关系查找元素
    
    Args:
        mapping: 元素映射字典
        element_index: 元素索引
        element_type: 元素类型
    
    Returns:
        tuple: (target_paragraph, element_info) 或 (None, None)
    """
    if element_index in mapping:
        element_info = mapping[element_index]
        if element_info.get("element_type") == element_type:
            return element_info.get("paragraph"), element_info
    
    return None, None
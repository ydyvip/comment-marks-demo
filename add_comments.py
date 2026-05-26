"""
docx_comment_demo - 为 DOCX 文档批量添加批注

批注 JSON 格式（支持段落和单元格的统一索引系统）：

[
  {
    "element_index": 2,          # 元素索引（段落或单元格，从 0 开始）
    "element_type": "paragraph", # 元素类型："paragraph" 或 "cell"
    "match_text": "工作成果",     # 在该元素内匹配的文本片段（第一次出现处）
    "match_occurrence": 1,       # 可选，同元素内第几次出现，默认 1
    "title": "标题",             # 批注标题（显示在批注正文前）
    "content": "批注内容"        # 批注正文
  }
]

优化特性：
- 统一索引系统：段落和单元格使用连续的索引编号
- 智能文本匹配：支持跨run的精确文本定位
- 单元格支持：可以直接在表格单元格内添加批注
- 便利函数：自动计算派生属性（has_text, is_empty, word_count等）

运行示例：
  python add_comments.py some.docx out.docx annotations.json --author 张三
  python add_comments.py some.docx --list-elements          # 查看元素结构
  python add_comments.py some.docx --document-json my_document.json  # 使用优化后的JSON结构
"""

import json
import sys
import tempfile
import xml.etree.ElementTree as ET
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent / "scripts"))
sys.path.insert(0, str(Path(__file__).parent / "scripts" / "office"))

from comment import add_comment
from office.pack import pack
from office.unpack import unpack

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


# ─────────────────────────────────────────────────────────────────
# 工具函数：解析 document.xml
# ─────────────────────────────────────────────────────────────────

def _parse_paragraphs(doc_xml_path: Path):
    """返回 (tree, root, paragraphs列表)"""
    tree = ET.parse(doc_xml_path)
    root = tree.getroot()
    body = root.find(f".//{{{W}}}body")
    paragraphs = body.findall(f".//{{{W}}}p") if body is not None else []
    return tree, root, paragraphs


def _para_style(para) -> str:
    pPr = para.find(f"{{{W}}}pPr")
    if pPr is None:
        return ""
    pStyle = pPr.find(f"{{{W}}}pStyle")
    return pStyle.get(f"{{{W}}}val", "") if pStyle is not None else ""


def _para_text(para) -> str:
    text = ""
    for run in para.findall(f".//{{{W}}}r"):
        t = run.find(f"{{{W}}}t")
        text += (t.text or "") if t is not None else ""
    return text


def _get_run_text(run) -> str:
    t = run.find(f"{{{W}}}t")
    return (t.text or "") if t is not None else ""


# ─────────────────────────────────────────────────────────────────
# 文档层级结构分析
# ─────────────────────────────────────────────────────────────────

def _get_document_hierarchy(paragraphs):
    """
    分析文档的层级结构，返回每个段落的标题级别信息
    
    返回：[(index, level, title, style), ...]
    level: 0-正文, 1-标题1, 2-标题2, 3-标题3, 4-标题4, ...
    """
    hierarchy = []
    STYLE_LEVEL_MAP = {
        "Title": 1,
        "Heading1": 1, "1": 1,
        "Heading2": 2, "2": 2,
        "Heading3": 3, "3": 3,
        "Heading4": 4, "4": 4,
        "Heading5": 5, "5": 5,
        "Heading6": 6, "6": 6,
    }
    
    for i, para in enumerate(paragraphs):
        style = _para_style(para)
        level = STYLE_LEVEL_MAP.get(style, 0)  # 0表示正文
        text = _para_text(para)
        
        hierarchy.append({
            "index": i,
            "level": level,
            "text": text,
            "style": style,
            "title": text[:30] if text else ""  # 标题文本前30字符
        })
    
    return hierarchy

def _find_paragraph_by_hierarchy(hierarchy, level=None, keyword=None, exact_level=False):
    """
    根据层级和关键词查找段落
    
    Args:
        hierarchy: 文档层级结构
        level: 标题级别（None表示不限制）
        keyword: 关键词（支持模糊匹配）
        exact_level: 是否要求精确层级匹配
    
    Returns: 符合条件的段落索引列表
    """
    matched = []
    
    for item in hierarchy:
        # 检查层级
        if level is not None:
            if exact_level:
                if item["level"] != level:
                    continue
            else:
                if item["level"] < level:  # 找到等于或大于指定层级的段落
                    continue
        
        # 检查关键词
        if keyword:
            text_lower = item["text"].lower()
            keyword_lower = keyword.lower()
            if keyword_lower not in text_lower:
                continue
        
        matched.append(item["index"])
    
    return matched


# ─────────────────────────────────────────────────────────────────
# 列出段落结构（供用户配置 JSON 时参考）
# ─────────────────────────────────────────────────────────────────

def list_elements(docx_path: str, document_json_path: str = None):
    """打印文档元素结构（段落和单元格），方便用户配置 element_index"""
    
    # 如果有文档JSON文件，显示优化后的结构
    if document_json_path and Path(document_json_path).exists():
        try:
            with open(document_json_path, 'r', encoding='utf-8') as f:
                doc_data = json.load(f)
            
            print(f"\n📊 优化后的文档结构（基于 {Path(document_json_path).name}）：")
            print(f"{'索引':>4}  {'类型':<8}  {'样式':<10}  {'位置':<20}  {'内容（前40字）'}")
            print("─" * 90)
            
            for element in doc_data.get("content", []):
                index = element.get("index", 0)
                element_type = element.get("element_type", "unknown")
                style = element.get("style", "其他")
                text = element.get("text", "")
                
                # 构建位置信息
                position = ""
                if element_type == "cell":
                    position = element.get("table_position", "N/A")
                else:
                    position = "段落"
                
                # 预览文本
                preview = text[:40].replace("\n", " ") + "..." if len(text) > 40 else text
                
                type_label = "段落" if element_type == "paragraph" else "单元格"
                print(f"  {index:>3}  {type_label:<8}  {style:<10}  {position:<20}  {preview}")
            
            print(f"\n💡 提示：使用 element_index 和 element_type 进行批注定位")
            print(f"     例如：{{\"element_index\": 5, \"element_type\": \"cell\", \"match_text\": \"内容\"}}")
            return
            
        except Exception as e:
            print(f"⚠️ 无法读取JSON文件: {e}")
    
    # 如果没有JSON文件，回退到段落列表
    with tempfile.TemporaryDirectory() as tmp:
        unpacked_dir = Path(tmp) / "unpacked"
        _, msg = unpack(docx_path, str(unpacked_dir))
        doc_xml = unpacked_dir / "word" / "document.xml"
        _, _, paragraphs = _parse_paragraphs(doc_xml)

    STYLE_LABEL = {
        "1": "标题1", "2": "标题2", "3": "标题3", "4": "标题4",
        "Heading1": "标题1", "Heading2": "标题2", "Heading3": "标题3",
        "Title": "标题",
    }

    print(f"\n📄 传统段落结构（仅段落，不包含单元格）：")
    print(f"{'索引':>4}  {'样式':<10}  {'段落内容（前50字）'}")
    print("─" * 72)
    for i, para in enumerate(paragraphs):
        style = _para_style(para)
        label = STYLE_LABEL.get(style, f"正文" if not style else style)
        text = _para_text(para)
        preview = text[:50].replace("\n", " ")
        print(f"  {i:>3}  {label:<10}  {preview}")
    print()
    print("💡 提示：推荐使用 --document-json 参数查看完整的段落和单元格结构")


# ─────────────────────────────────────────────────────────────────
# 核心：在指定段落内定位 match_text，找到对应的 run 元素（改进版）
# ─────────────────────────────────────────────────────────────────

def _normalize_text(text: str) -> str:
    """标准化文本，用于匹配比较"""
    import re
    # 移除多余空格，换行符转为空格，统一为中文标点
    normalized = re.sub(r'\s+', ' ', text.strip())
    normalized = normalized.replace('\n', ' ')
    # 标准化中文标点
    normalized = normalized.replace('，', ',').replace('。', '.').replace('？', '?').replace('！', '!')
    normalized = normalized.replace('：', ':').replace('；', ';').replace('"', '"').replace('"', '"')
    normalized = normalized.replace(''', "'").replace(''', "'")
    # 转换为小写进行匹配（可选）
    # normalized = normalized.lower()
    return normalized

def _locate_run_in_para(para, match_text: str, occurrence: int = 1):
    """
    在段落内找第 occurrence 次出现 match_text 的位置，返回对应的 run 元素。
    完全重写版：确保精确的文本匹配和run定位。

    策略：
    1. 直接匹配段落中的run文本
    2. 支持跨run的文本匹配
    3. 确保定位的精确性

    返回 (run_elem, found, matched_text) 或 (None, False, None)
    """
    import re
    
    # 获取段落所有run
    runs = para.findall(f".//{{{W}}}r")
    if not runs:
        return None, False, None
    
    # 构建run文本映射
    run_texts = []
    for run in runs:
        t = run.find(f"{{{W}}}t")
        if t is not None and t.text:
            run_texts.append(t.text.strip())
        else:
            run_texts.append("")
    
    # 拼接完整段落文本
    full_text = "".join(run_texts)
    
    # 如果匹配文本为空，返回第一个非空run
    if not match_text.strip():
        for run in runs:
            t = run.find(f"{{{W}}}t")
            if t is not None and t.text and t.text.strip():
                return run, True, t.text.strip()
        return None, False, None
    
    # 标准化文本（不转换为小写，保持原样）
    def normalize_simple(text):
        import re
        # 移除多余空格，换行符转为空格
        normalized = re.sub(r'\s+', ' ', text.strip())
        normalized = normalized.replace('\n', ' ')
        return normalized
    
    normalized_match = normalize_simple(match_text)
    normalized_full = normalize_simple(full_text)
    
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
    
    # 找到覆盖目标位置的run
    current_pos = 0
    for i, run_text in enumerate(run_texts):
        if not run_text:
            continue
        
        run_start = current_pos
        run_end = current_pos + len(run_text)
        
        # 检查目标位置是否在这个run中
        if run_start <= target_pos < run_end:
            target_run = runs[i]
            matched_text = normalized_match
            return target_run, True, matched_text
        
        current_pos = run_end
    
    # 如果跨多个run，返回第一个包含匹配文本开头的run
    if target_pos < len(full_text):
        for i, run_text in enumerate(run_texts):
            if not run_text:
                continue
            
            run_start = current_pos
            run_end = current_pos + len(run_text)
            
            # 如果匹配文本跨多个run，返回第一个run
            if run_start <= target_pos < run_end:
                target_run = runs[i]
                matched_text = normalized_match
                return target_run, True, matched_text
            
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


# ─────────────────────────────────────────────────────────────────
# 在 document.xml 注入批注标记（字符串操作）
# ─────────────────────────────────────────────────────────────────

def _inject_comment_markers(doc_xml_path: Path, comment_id: int, target_run) -> bool:
    """
    在 target_run 前后插入 commentRangeStart/End 和 commentReference。
    通过 run 中的文本内容在 pretty-printed XML 中定位并注入。
    """
    xml_text = doc_xml_path.read_text(encoding="utf-8")
    run_text = _get_run_text(target_run)

    if not run_text:
        print(f"  [WARN] comment {comment_id}: 目标 run 文本为空，跳过标记注入")
        return False

    # 尝试找 <w:t>文本</w:t>，支持带/不带 xml:space 属性
    t_pos = -1
    for pattern in [f"<w:t>{run_text}</w:t>", f'<w:t xml:space="preserve">{run_text}</w:t>']:
        t_pos = xml_text.find(pattern)
        if t_pos != -1:
            break

    if t_pos == -1:
        print(f"  [WARN] comment {comment_id}: 找不到文本 '{run_text[:20]}'，跳过标记注入")
        return False

    run_start = xml_text.rfind("<w:r", 0, t_pos)
    if run_start == -1:
        return False
    run_end = xml_text.find("</w:r>", t_pos)
    if run_end == -1:
        return False
    run_end += len("</w:r>")

    run_fragment = xml_text[run_start:run_end]
    start_tag = f'<w:commentRangeStart w:id="{comment_id}"/>'
    end_tag = (
        f'<w:commentRangeEnd w:id="{comment_id}"/>'
        f'<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr>'
        f'<w:commentReference w:id="{comment_id}"/></w:r>'
    )

    xml_text = (
        xml_text[:run_start]
        + start_tag
        + run_fragment
        + end_tag
        + xml_text[run_end:]
    )
    doc_xml_path.write_text(xml_text, encoding="utf-8")
    return True


# ─────────────────────────────────────────────────────────────────
# 主流程
# ─────────────────────────────────────────────────────────────────

def add_comments_from_json(
    input_docx: str,
    output_docx: str,
    annotations: list,
    author: str = "Reviewer",
    document_json_path: str = None,
):
    """
    参数
    ----
    input_docx       : 原始 docx 路径
    output_docx      : 输出 docx 路径
    annotations      : 批注对象列表，每项包含：
                       element_index      元素索引（必填）
                       element_type       元素类型："paragraph" 或 "cell"（必填）
                       match_text         元素内匹配文本（必填）
                       match_occurrence   同元素内第几次出现，默认 1
                       title              批注标题
                       content            批注正文
    author           : 批注作者
    document_json_path: 优化后的JSON文档路径（可选，用于获取元素信息）
    """
    input_path = Path(input_docx)
    if not input_path.exists():
        raise FileNotFoundError(f"找不到文件: {input_docx}")

    # 加载文档JSON数据（如果提供）
    doc_data = None
    if document_json_path and Path(document_json_path).exists():
        try:
            with open(document_json_path, 'r', encoding='utf-8') as f:
                doc_data = json.load(f)
            print(f"[1/4] 加载文档结构: {Path(document_json_path).name}")
        except Exception as e:
            print(f"⚠️ 无法加载JSON文件: {e}，将回退到原始处理方式")

    with tempfile.TemporaryDirectory() as tmp:
        unpacked_dir = Path(tmp) / "unpacked"

        # Step 1: 解包
        print(f"[1/4] 解包 {input_docx} ...")
        _, msg = unpack(str(input_path), str(unpacked_dir))
        print(f"      {msg}")

        doc_xml = unpacked_dir / "word" / "document.xml"
        _, _, paragraphs = _parse_paragraphs(doc_xml)
        
        if doc_data:
            total_elements = len(doc_data.get("content", []))
            print(f"[2/4] 文档共 {total_elements} 个元素（段落+单元格）")
            
            # 分析文档层级结构
            hierarchy = _get_document_hierarchy(paragraphs)
            print(f"      文档层级：共 {len([h for h in hierarchy if h['level'] > 0])} 个标题级别段落")
        else:
            print(f"[2/4] 文档共 {len(paragraphs)} 个段落（原始模式）")

        # Step 3: 逐条插入批注
        print(f"[3/4] 插入 {len(annotations)} 条批注 ...")
        success_count = 0

        for cid, ann in enumerate(annotations):
            element_index = ann.get("element_index")
            element_type = ann.get("element_type", "paragraph")
            match_text = ann.get("match_text", "")
            occurrence = ann.get("match_occurrence", 1)
            title = ann.get("title", f"Comment {cid}")
            content = ann.get("content", "")

            # 校验
            if element_index is None:
                print(f"  [SKIP] comment {cid}: 缺少 element_index")
                continue
            
            if element_type not in ["paragraph", "cell"]:
                print(f"  [SKIP] comment {cid}: element_type 必须是 'paragraph' 或 'cell'")
                continue
                
            if not match_text:
                print(f"  [SKIP] comment {cid}: 缺少 match_text")
                continue

            # 根据索引和类型查找目标元素
            target_element = None
            target_paragraph = None
            element_info = ""
            
            if doc_data:
                # 使用优化的JSON数据进行查找
                target_element = next((e for e in doc_data.get("content", []) 
                                     if e.get("index") == element_index and e.get("element_type") == element_type), None)
                
                if target_element:
                    element_info = f"[{target_element.get('style', '其他')}]"
                    if element_type == "cell":
                        element_info += f" {target_element.get('table_position', 'N/A')}"
                    else:
                        element_info += f" 段落内容预览: {target_element.get('text', '')[:30]}"
                
                if not target_element:
                    print(f"  [SKIP] comment {cid}: 找不到索引={element_index}, 类型={element_type} 的元素")
                    continue
                
                # 在文档XML中找到对应的段落
                if element_type == "paragraph":
                    if element_index < len(paragraphs):
                        target_paragraph = paragraphs[element_index]
                    else:
                        print(f"  [SKIP] comment {cid}: 段落索引超出范围")
                        continue
                        
            else:
                # 回退到原始模式
                if element_type == "paragraph":
                    if element_index >= len(paragraphs) or element_index < 0:
                        print(f"  [SKIP] comment {cid}: 段落索引={element_index} 超出范围")
                        continue
                    target_paragraph = paragraphs[element_index]
                    element_info = f"段落[{element_index}]"
                else:
                    print(f"  [SKIP] comment {cid}: 原始模式不支持单元格批注")
                    continue

            # 显示元素信息
            if doc_data:
                print(f"  [INFO] comment {cid}: {element_info}")

            # 定位 run（改进版）
            if target_paragraph:
                target_run, found, matched_text = _locate_run_in_para(target_paragraph, match_text, occurrence)
                if not found or target_run is None:
                    print(f"  [WARN] comment {cid}: 元素中找不到 '{match_text}'，尝试在相邻元素搜索...")
                    found_in_adjacent = False
                    
                    if doc_data:
                        # 在同类型的相邻元素中搜索
                        adjacent_elements = [e for e in doc_data.get("content", []) 
                                           if e.get("element_type") == element_type]
                        
                        for adj_elem in adjacent_elements:
                            adj_index = adj_elem.get("index")
                            if adj_index != element_index and adj_index < len(paragraphs):
                                adj_para = paragraphs[adj_index]
                                target_run, found, matched_text = _locate_run_in_para(adj_para, match_text, occurrence)
                                if found and target_run is not None:
                                    target_paragraph = adj_para
                                    element_index = adj_index
                                    found_in_adjacent = True
                                    print(f"  [FOUND] comment {cid}: 在相邻元素[{adj_index}] 找到匹配")
                                    break
                        
                        if not found_in_adjacent:
                            print(f"  [SKIP] comment {cid}: 所有{element_type}中都找不到第{occurrence}次出现的 '{match_text}'")
                            continue
                    else:
                        print(f"  [SKIP] comment {cid}: 原始模式中找不到匹配文本")
                        continue

            # 写入 comments.xml
            comment_text = f"{title}: {content}"
            _, result_msg = add_comment(
                str(unpacked_dir),
                comment_id=cid,
                text=comment_text,
                author=author,
                initials=author[0].upper() if author else "R",
            )

            # 注入 document.xml 标记
            ok = _inject_comment_markers(doc_xml, cid, target_run)
            if ok:
                run_text_preview = _get_run_text(target_run)[:15]
                element_text = f"{element_type}[{element_index}]"
                print(f"  [✓] comment {cid}: {element_text} '{matched_text}' → run='{run_text_preview}'")
                success_count += 1
            else:
                print(f"  [✗] comment {cid}: XML 标记注入失败")

        # Step 4: 重新打包
        print(f"[4/4] 打包输出 {output_docx} ...")
        _, pack_msg = pack(str(unpacked_dir), output_docx, original_file=str(input_path))
        print(f"      {pack_msg}")

    print(f"\n✅ 完成！成功插入 {success_count}/{len(annotations)} 条批注 → {output_docx}")


# ─────────────────────────────────────────────────────────────────
# 命令行入口
# ─────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(
        description="根据 JSON 批注数据为 DOCX 文档精确添加批注",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
定位方式：支持段落和单元格的统一索引系统，彻底避免重复内容误匹配。

使用步骤：
  1. 先查看元素结构，确定 element_index：
        python add_comments.py some.docx --list-elements
        python add_comments.py some.docx --list-elements --document-json my_document.json

  2. 编写批注 JSON（annotations.json）：
        [
          {
            "element_index": 1,
            "element_type": "paragraph",
            "match_text": "工作成果",
            "match_occurrence": 1,
            "title": "格式问题",
            "content": "建议使用列表展示"
          },
          {
            "element_index": 985,
            "element_type": "cell",
            "match_text": "项目数据",
            "title": "数据检查",
            "content": "需要补充完整"
          }
        ]

  3. 执行添加：
        python add_comments.py some.docx output.docx annotations.json --author 张三
        python add_comments.py some.docx output.docx annotations.json --document-json my_document.json --author 张三
""",
    )
    parser.add_argument("input_docx", help="输入 docx 文件路径")
    parser.add_argument("output_docx", nargs="?", help="输出 docx 文件路径（--list-elements 时可省略）")
    parser.add_argument("annotations_json", nargs="?", help="批注 JSON 文件路径或内联 JSON 字符串")
    parser.add_argument("--author", default="Reviewer", help="批注作者姓名（默认 Reviewer）")
    parser.add_argument("--list-paragraphs", action="store_true", help="列出文档段落结构后退出（兼容旧版本）")
    parser.add_argument("--list-elements", action="store_true", help="列出文档元素结构后退出")
    parser.add_argument("--document-json", help="优化后的文档JSON文件路径（用于查看完整结构）")

    args = parser.parse_args()

    if args.list_paragraphs:
        list_paragraphs(args.input_docx)
        sys.exit(0)
    
    if args.list_elements:
        list_elements(args.input_docx, args.document_json)
        sys.exit(0)

    if not args.output_docx or not args.annotations_json:
        parser.error("添加批注时需要提供 output_docx 和 annotations_json")

    # 解析批注数据
    ann_input = args.annotations_json
    try:
        ann_path = Path(ann_input)
        annotations = json.loads(ann_path.read_text(encoding="utf-8") if ann_path.exists() else ann_input)
    except (json.JSONDecodeError, OSError) as e:
        print(f"❌ 无法解析批注数据: {e}")
        sys.exit(1)

    if not isinstance(annotations, list):
        print("❌ 批注数据必须是 JSON 数组格式")
        sys.exit(1)

    add_comments_from_json(
        input_docx=args.input_docx,
        output_docx=args.output_docx,
        annotations=annotations,
        author=args.author,
        document_json_path=args.document_json,
    )

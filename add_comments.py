"""
docx_comment_demo - 为 DOCX 文档批量添加批注

批注 JSON 格式（每条批注精确锚定到"段落 + 段内文本"，避免重复内容误匹配）：

[
  {
    "para_index":  2,           # 段落索引（从 0 开始，可用 list_paragraphs() 查看）
    "match_text":  "工作成果",   # 在该段落内匹配的文本片段（第一次出现处）
    "match_occurrence": 1,      # 可选，同段内第几次出现，默认 1
    "title":       "标题",       # 批注标题（显示在批注正文前）
    "content":     "批注内容"    # 批注正文
  }
]

运行示例：
  python add_comments.py some.docx out.docx annotations.json --author 张三
  python add_comments.py some.docx --list-paragraphs          # 查看段落结构
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
# 列出段落结构（供用户配置 JSON 时参考）
# ─────────────────────────────────────────────────────────────────

def list_paragraphs(docx_path: str):
    """打印文档段落结构，方便用户配置 para_index"""
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

    print(f"\n{'索引':>4}  {'样式':<10}  {'段落内容（前50字）'}")
    print("─" * 72)
    for i, para in enumerate(paragraphs):
        style = _para_style(para)
        label = STYLE_LABEL.get(style, f"正文" if not style else style)
        text = _para_text(para)
        preview = text[:50].replace("\n", " ")
        print(f"  {i:>3}  {label:<10}  {preview}")
    print()


# ─────────────────────────────────────────────────────────────────
# 核心：在指定段落内定位 match_text，找到对应的 run 元素
# ─────────────────────────────────────────────────────────────────

def _locate_run_in_para(para, match_text: str, occurrence: int = 1):
    """
    在段落内找第 occurrence 次出现 match_text 的位置，返回对应的 run 元素。

    策略：
    1. 先拼出段落全文，找到 match_text 在段落内的字符偏移（第 occurrence 次）
    2. 遍历 run，累计字符数，找到覆盖该偏移的 run

    返回 (run_elem, found) 或 (None, False)
    """
    runs = para.findall(f".//{{{W}}}r")
    full_text = ""
    run_spans = []  # (start, end, run_elem)

    for run in runs:
        t = run.find(f"{{{W}}}t")
        text = (t.text or "") if t is not None else ""
        start = len(full_text)
        full_text += text
        if text:
            run_spans.append((start, len(full_text), run))

    # 找第 occurrence 次出现
    pos = -1
    count = 0
    search_from = 0
    while count < occurrence:
        pos = full_text.find(match_text, search_from)
        if pos == -1:
            return None, False
        count += 1
        search_from = pos + 1

    # 找覆盖 pos 的 run
    for start, end, run in run_spans:
        if start <= pos < end:
            return run, True

    return None, False


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
):
    """
    参数
    ----
    input_docx   : 原始 docx 路径
    output_docx  : 输出 docx 路径
    annotations  : 批注对象列表，每项包含：
                     para_index       段落索引（必填）
                     match_text       段内匹配文本（必填）
                     match_occurrence 同段内第几次出现，默认 1
                     title            批注标题
                     content          批注正文
    author       : 批注作者
    """
    input_path = Path(input_docx)
    if not input_path.exists():
        raise FileNotFoundError(f"找不到文件: {input_docx}")

    with tempfile.TemporaryDirectory() as tmp:
        unpacked_dir = Path(tmp) / "unpacked"

        # Step 1: 解包
        print(f"[1/4] 解包 {input_docx} ...")
        _, msg = unpack(str(input_path), str(unpacked_dir))
        print(f"      {msg}")

        doc_xml = unpacked_dir / "word" / "document.xml"
        _, _, paragraphs = _parse_paragraphs(doc_xml)
        print(f"[2/4] 文档共 {len(paragraphs)} 个段落")

        # Step 3: 逐条插入批注
        print(f"[3/4] 插入 {len(annotations)} 条批注 ...")
        success_count = 0

        for cid, ann in enumerate(annotations):
            para_index = ann.get("para_index")
            match_text = ann.get("match_text", "")
            occurrence = ann.get("match_occurrence", 1)
            title = ann.get("title", f"Comment {cid}")
            content = ann.get("content", "")

            # 校验
            if para_index is None:
                print(f"  [SKIP] comment {cid}: 缺少 para_index")
                continue
            if para_index >= len(paragraphs) or para_index < 0:
                print(f"  [SKIP] comment {cid}: para_index={para_index} 超出范围（共 {len(paragraphs)} 段）")
                continue
            if not match_text:
                print(f"  [SKIP] comment {cid}: 缺少 match_text")
                continue

            para = paragraphs[para_index]
            para_full_text = _para_text(para)

            # 定位 run
            target_run, found = _locate_run_in_para(para, match_text, occurrence)
            if not found or target_run is None:
                print(f"  [SKIP] comment {cid}: 段落[{para_index}] 中找不到第{occurrence}次出现的 '{match_text}'")
                print(f"         段落全文: '{para_full_text[:60]}'")
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
                print(f"  [✓] comment {cid}: 段落[{para_index}] 第{occurrence}次 '{match_text}' → run='{run_text_preview}'")
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
定位方式：段落索引 + 段内文本匹配，彻底避免重复内容误匹配。

使用步骤：
  1. 先查看段落结构，确定 para_index：
       python add_comments.py some.docx --list-paragraphs

  2. 编写批注 JSON（annotations.json）：
       [
         {
           "para_index": 1,
           "match_text": "工作成果",
           "match_occurrence": 1,
           "title": "格式问题",
           "content": "建议使用列表展示"
         }
       ]

  3. 执行添加：
       python add_comments.py some.docx output.docx annotations.json --author 张三
""",
    )
    parser.add_argument("input_docx", help="输入 docx 文件路径")
    parser.add_argument("output_docx", nargs="?", help="输出 docx 文件路径（--list-paragraphs 时可省略）")
    parser.add_argument("annotations_json", nargs="?", help="批注 JSON 文件路径或内联 JSON 字符串")
    parser.add_argument("--author", default="Reviewer", help="批注作者姓名（默认 Reviewer）")
    parser.add_argument("--list-paragraphs", action="store_true", help="列出文档段落结构后退出")

    args = parser.parse_args()

    if args.list_paragraphs:
        list_paragraphs(args.input_docx)
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
    )

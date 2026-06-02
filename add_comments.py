"""
add_comments.py - 为 DOCX 文档批量添加批注

基于 docx-comments 库，支持精准文本锚定批注。

批注 JSON 格式：
[
  {
    "match_text": "工作成果",     # 全文匹配的文本片段
    "match_occurrence": 1,       # 可选，全篇第几次出现，默认 1
    "title": "格式问题",          # 批注标题
    "content": "建议使用列表展示"  # 批注正文
  }
]

使用示例：
  python add_comments.py input.docx output.docx annotations.json --author 张三
  python add_comments.py input.docx --list-paragraphs
"""

import json
import sys
from pathlib import Path

from docx import Document
from docx_comments import CommentManager, PersonInfo


def _get_run_text(run) -> str:
    """获取单个 run 的文本内容"""
    t = run.find("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t")
    return (t.text or "") if t is not None else ""


def _get_para_text(paragraph) -> str:
    """从 lxml 段落元素中提取完整文本（拼接所有 run）"""
    return "".join(_get_run_text(run) for run in paragraph.findall(
        ".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r"
    ))


def _get_para_style(paragraph) -> str:
    """获取段落样式名"""
    pPr = paragraph.find(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pPr")
    if pPr is None:
        return ""
    pStyle = pPr.find(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pStyle")
    return pStyle.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val", "") if pStyle is not None else ""


def _find_text_runs(paragraph, match_text: str):
    """
    在段落内定位 match_text 的位置，
    返回 (start_run_idx, end_run_idx) 用于锚定批注。

    通过累加各 run 的字符位置来计算。
    """
    W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    runs = paragraph.findall(f".//{W}r")
    run_texts = []
    for run in runs:
        text = _get_run_text(run)
        run_texts.append(text)

    full_text = "".join(run_texts)

    if not match_text:
        # 空匹配文本 -> 使用第一个非空 run
        for i, rt in enumerate(run_texts):
            if rt.strip():
                return (i, i)
        return (0, 0) if runs else (0, 0)

    # 查找 match_text 在段落中的位置
    pos = full_text.find(match_text)
    if pos == -1:
        return None, None

    match_start = pos
    match_end = pos + len(match_text)

    # 计算 match_start 和 match_end 落在哪个 run 中
    char_offset = 0
    start_run_idx = 0
    end_run_idx = 0
    found_start = False

    for i, rt in enumerate(run_texts):
        run_start = char_offset
        run_end = char_offset + len(rt)
        if rt:
            if not found_start and run_start <= match_start < run_end:
                start_run_idx = i
                end_run_idx = i
                found_start = True
            elif found_start and run_start <= match_end <= run_end:
                end_run_idx = i
                break
            elif found_start and match_end > run_end:
                end_run_idx = i
        char_offset = run_end

    return start_run_idx, end_run_idx


def list_paragraphs(docx_path: str):
    """列出文档所有段落及其样式"""
    doc = Document(docx_path)

    STYLE_LABEL = {
        "Title": "标题",
        "Heading1": "标题1", "Heading2": "标题2", "Heading3": "标题3",
        "Heading4": "标题4", "Heading5": "标题5", "Heading6": "标题6",
        "Normal": "正文",
    }

    print(f"\n📄 段落列表（共 {len(doc.paragraphs)} 段）：")
    print(f"{'索引':>5}  {'样式':<12}  {'内容（前60字）'}")
    print("─" * 80)

    for i, para in enumerate(doc.paragraphs):
        style = para.style.name if para.style else ""
        label = STYLE_LABEL.get(style, style if style else "正文")
        text = para.text[:60] + "..." if len(para.text) > 60 else para.text
        print(f"  {i:>4}  {label:<12}  {text}")

    print(f"\n💡 使用 --list-paragraphs 查看段落，配合 --author 参数添加批注")


def add_comments_from_json(
    input_docx: str,
    output_docx: str,
    annotations: list,
    author: str = "Reviewer",
):
    """
    基于 docx-comments 库，为 DOCX 文档添加批注。

    参数
    ----
    input_docx       : 输入 docx 路径
    output_docx      : 输出 docx 路径
    annotations      : 批注列表
    author           : 批注作者
    """
    input_path = Path(input_docx)
    if not input_path.exists():
        raise FileNotFoundError(f"找不到文件: {input_docx}")

    print(f"加载文档: {input_path.name}")
    doc = Document(str(input_path))

    mgr = CommentManager(doc)

    # 确保作者存在于 people.xml 中
    try:
        person = mgr.get_person(author)
    except KeyError:
        person = mgr.ensure_person(author)

    initials = author[0].upper() if author else "R"

    total = len(annotations)
    success_count = 0

    print(f"准备插入 {total} 条批注 ...")

    for cid, ann in enumerate(annotations):
        match_text = ann.get("match_text", "")
        occurrence = ann.get("match_occurrence", 1)
        title = ann.get("title", f"Comment {cid}")
        content = ann.get("content", "")

        # 校验必填字段
        if not match_text:
            print(f"  [SKIP] comment {cid}: 缺少 match_text")
            continue

        # 全局搜索包含 match_text 的段落，定位第 occurrence 次出现
        para_idx = None
        target_para = None
        para_text = None
        found_count = 0

        for idx, para in enumerate(doc.paragraphs):
            text = para.text
            if match_text in text:
                found_count += 1
                if found_count == occurrence:
                    para_idx = idx
                    target_para = para
                    para_text = text
                    break

        if target_para is None:
            print(f"  [SKIP] comment {cid}: 全文中找不到第 {occurrence} 次出现的 '{match_text}'")
            continue

        # 对于 lxml 段落元素，需要用 XML 方式获取 runs
        # python-docx Paragraph 的 .element 属性是 lxml Element
        para_elem = target_para._element if hasattr(target_para, '_element') else target_para.element

        # 计算 run 索引范围用于锚定
        start_run, end_run = _find_text_runs(para_elem, match_text)

        if start_run is None:
            print(f"  [SKIP] comment {cid}: 无法定位文本 '{match_text}' 的 run 位置")
            continue

        # 构建批注内容
        comment_text = f"{title}: {content}" if title and content else (title or content)

        try:
            comment_id = mgr.add_comment(
                paragraph=target_para,
                start_run=start_run,
                end_run=end_run,
                text=comment_text,
                author=person,
                initials=initials,
            )
            print(f"  [✓] comment {cid}: [{para_idx}] '{match_text}' → {author}")
            success_count += 1
        except Exception as e:
            print(f"  [✗] comment {cid}: 添加失败 - {e}")

    # 保存文档
    print(f"\n保存输出: {output_docx}")
    doc.save(output_docx)

    print(f"\n✅ 完成！成功插入 {success_count}/{total} 条批注 → {output_docx}")


def main():
    import argparse

    parser = argparse.ArgumentParser(
        description="基于 docx-comments 为 DOCX 文档批量添加批注",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
使用示例：
  # 列出文档段落
  python add_comments.py document.docx --list-paragraphs

  # 添加批注（从 JSON 文件）
  python add_comments.py document.docx output.docx annotations.json --author 张三

  # 添加批注（从内联 JSON）
  python add_comments.py document.docx output.docx '[{"match_text":"文本","title":"批注","content":"内容"}]'
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
        annotations = json.loads(
            ann_path.read_text(encoding="utf-8") if ann_path.exists() else ann_input
        )
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


if __name__ == "__main__":
    main()

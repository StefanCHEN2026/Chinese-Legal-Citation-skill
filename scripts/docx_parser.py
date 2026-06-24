#!/usr/bin/env python3
"""
法学论文脚注解析模块
从 .docx 文件中提取所有脚注内容，输出 JSON 格式供核查使用。

依赖：pip install python-docx lxml
"""

import json
import zipfile
import sys
import argparse
import io
from pathlib import Path
from io import BytesIO

try:
    from lxml import etree
except ImportError:
    print("错误：需要安装 lxml 库。请运行: pip install lxml")
    sys.exit(1)

try:
    from docx import Document
except ImportError:
    print("错误：需要安装 python-docx 库。请运行: pip install python-docx")
    sys.exit(1)

# OOXML namespace map
NSMAP = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
    'w14': 'http://schemas.microsoft.com/office/word/2010/wordml',
}


def extract_text_from_element(element):
    """从 OOXML 元素中递归提取纯文本"""
    texts = []
    # 提取 w:t 元素中的文本
    for t in element.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t'):
        if t.text:
            texts.append(t.text)
    # 处理 delText（修订模式中的删除文本），不提取
    return ''.join(texts)


def extract_footnote_text_without_deletions(element):
    """提取脚注文本，跳过 w:del 中的删除文本"""
    texts = []
    # 获取所有段落
    for para in element.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p'):
        # 跳过删除修订中的段落
        pPr = para.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pPr')
        if pPr is not None:
            rPr_del = pPr.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rPr')
            if rPr_del is not None and rPr_del.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}del') is not None:
                continue

        para_text = []
        for r in para.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r'):
            # 跳过删除修订中的 run
            rPr = r.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rPr')
            is_deleted = False
            if rPr is not None:
                if rPr.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}del') is not None:
                    is_deleted = True
            if is_deleted:
                continue

            # 跳过脚注引用标记
            footnote_ref = r.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}footnoteRef')
            if footnote_ref is not None:
                continue

            t = r.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t')
            if t is not None and t.text:
                para_text.append(t.text)

        texts.append(''.join(para_text))

    return ''.join(texts)


def get_paragraph_index_for_footnote(docx_path, footnote_id):
    """尝试找到脚注引用在正文中的位置"""
    with zipfile.ZipFile(docx_path, 'r') as z:
        document_xml = z.read('word/document.xml')

    tree = etree.fromstring(document_xml)

    para_index = -1
    for i, para in enumerate(tree.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p')):
        for ref in para.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}footnoteReference'):
            if ref.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id') == str(footnote_id):
                para_index = i
                break
        if para_index >= 0:
            break

    return para_index


def parse_footnotes(docx_path):
    """
    解析 docx 文件中的所有脚注。

    返回：
        list[dict]: 脚注列表，每个脚注包含 id, text, paragraph_index
    """
    footnotes = []

    try:
        with zipfile.ZipFile(docx_path, 'r') as z:
            # 检查是否包含脚注
            if 'word/footnotes.xml' not in z.namelist():
                print("注意：文档中未找到脚注（无 word/footnotes.xml）。")
                return footnotes

            footnotes_xml = z.read('word/footnotes.xml')
    except FileNotFoundError:
        print(f"错误：文件不存在: {docx_path}")
        sys.exit(1)
    except zipfile.BadZipFile:
        print(f"错误：文件不是有效的 .docx 文件: {docx_path}")
        sys.exit(1)

    tree = etree.fromstring(footnotes_xml)

    for fn in tree.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}footnote'):
        fn_id = fn.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id')

        # 跳过特殊脚注：id=0 是分隔符脚注，id=-1 是分隔符
        if fn_id in ('0', '-1'):
            continue

        text = extract_footnote_text_without_deletions(fn)

        # 清理文本：去除多余空白
        text = text.strip()

        if not text:
            continue

        # 查找正文中的引用位置
        para_index = get_paragraph_index_for_footnote(docx_path, fn_id)

        # 将 XML 元素序列化为字符串供后续修订使用
        xml_str = etree.tostring(fn, encoding='unicode')

        footnotes.append({
            'id': fn_id,
            'text': text,
            'paragraph_index': para_index,
            'xml_structure': xml_str,
        })

    return footnotes


def main():
    # 修复 Windows 控制台编码问题
    if sys.platform == 'win32':
        try:
            sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
        except (AttributeError, OSError):
            pass

    parser = argparse.ArgumentParser(description='解析 .docx 文件中的脚注')
    parser.add_argument('input', type=str, help='输入的 .docx 文件路径')
    parser.add_argument('--output', '-o', type=str, default=None, help='输出的 JSON 文件路径（默认输出到标准输出）')
    parser.add_argument('--plain', action='store_true', help='只输出纯文本列表')

    args = parser.parse_args()

    footnotes = parse_footnotes(args.input)

    if args.plain:
        # 只输出文本
        for fn in footnotes:
            print(f"[脚注 {fn['id']}] (段落 {fn['paragraph_index']})")
            print(fn['text'])
            print()
    else:
        # 输出 JSON
        output = json.dumps(footnotes, ensure_ascii=False, indent=2)
        if args.output:
            Path(args.output).write_text(output, encoding='utf-8')
            print(f"已输出到: {args.output}")
        else:
            print(output)

    print(f"\n共找到 {len(footnotes)} 个脚注。")


if __name__ == '__main__':
    main()

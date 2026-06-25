#!/usr/bin/env python3
"""
法学论文脚注修订模块
对 .docx 文件中的脚注执行修订模式（Track Changes）修改和蓝色批注。

修订方式：
- 删除内容：用 <w:del> 包裹旧文本
- 插入内容：用 <w:ins> 包裹新文本
- 蓝色批注：修改对应run的颜色属性为蓝色

依赖：pip install python-docx lxml
"""

import json
import zipfile
import sys
import argparse
import shutil
import datetime
import copy
from pathlib import Path
from io import BytesIO

try:
    from lxml import etree
except ImportError:
    print("错误：需要安装 lxml 库。请运行: pip install lxml")
    sys.exit(1)

try:
    from docx import Document
    from docx.oxml.ns import qn, nsdecls
    from docx.oxml import parse_xml
except ImportError:
    print("错误：需要安装 python-docx 库。请运行: pip install python-docx")
    sys.exit(1)

# OOXML namespace
WML_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

# 修订作者信息
REVISION_AUTHOR = "法学引注核查"
REVISION_DATE = datetime.datetime.now().strftime("%Y-%m-%dT%H:%M:%SZ")
REVISION_ID_BASE = 1000  # 修订ID起始值

NSMAP = {
    'w': WML_NS,
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
}

# 默认字体配置
FONT_SIMSUN = "宋体"
FONT_TIMES_NEW_ROMAN = "Times New Roman"


def detect_text_language(text):
    """
    检测文本语言：返回 'zh'（中文为主）或 'en'（英文为主）。

    规则：只要包含任意一个CJK字符（中日韩统一表意文字），就判定为中文。
    """
    if not text:
        return 'zh'
    for ch in text:
        cp = ord(ch)
        # CJK统一表意文字 U+4E00-U+9FFF
        # CJK扩展A U+3400-U+4DBF
        # CJK兼容表意文字 U+F900-U+FAFF
        # 全角符号
        if (0x4E00 <= cp <= 0x9FFF or
            0x3400 <= cp <= 0x4DBF or
            0xF900 <= cp <= 0xFAFF or
            0xFF00 <= cp <= 0xFFEF):
            return 'zh'
    return 'en'


def get_default_font_props(lang):
    """
    根据语言返回默认字体属性。

    参数：
        lang: 'zh' 或 'en'

    返回：
        dict: {'ascii': str, 'hAnsi': str, 'eastAsia': str, 'sz': int | None}
    """
    if lang == 'zh':
        return {
            'ascii': FONT_TIMES_NEW_ROMAN,
            'hAnsi': FONT_SIMSUN,        # 中文上下文中的标点符号（引号等）使用宋体
            'eastAsia': FONT_SIMSUN,
            'sz': None  # 将在调用方从原始run获取
        }
    else:
        return {
            'ascii': FONT_TIMES_NEW_ROMAN,
            'hAnsi': FONT_TIMES_NEW_ROMAN,
            'eastAsia': FONT_TIMES_NEW_ROMAN,
            'sz': None
        }


def extract_font_props_from_run(run_elem):
    """
    从现有 w:r 元素中提取字体属性（rFonts 和 sz）。

    支持 Word 和 WPS 生成的 docx：
    - Word：从 w:rPr/w:rFonts 直接提取
    - WPS：若无 rFonts 而有 w:rStyle，记录样式引用后回退到默认字体
    - 默认字体（SimSun/TNR）确保 WPS 文档也能正确渲染

    参数：
        run_elem: lxml Element (w:r)

    返回：
        dict: {'ascii': str|None, 'hAnsi': str|None, 'eastAsia': str|None, 'sz': str|None, 'rStyle': str|None, 'italic': bool|None}
    """
    props = {'ascii': None, 'hAnsi': None, 'eastAsia': None, 'sz': None, 'rStyle': None, 'italic': None}
    rPr = run_elem.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rPr')
    if rPr is None:
        return props
    rFonts = rPr.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rFonts')
    if rFonts is not None:
        props['ascii'] = rFonts.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}ascii')
        props['hAnsi'] = rFonts.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}hAnsi')
        props['eastAsia'] = rFonts.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}eastAsia')
    else:
        # 检查 WPS 风格：通过 rStyle 引用样式
        rStyle = rPr.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rStyle')
        if rStyle is not None:
            props['rStyle'] = rStyle.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val')
    sz = rPr.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sz')
    if sz is not None:
        props['sz'] = sz.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val')
    # 检测斜体状态
    i_elem = rPr.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}i')
    if i_elem is not None:
        val = i_elem.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val')
        # val == 'false' or '0' 表示关闭斜体；无 val 或 val='true' 表示开启
        props['italic'] = not (val == 'false' or val == '0')
    return props


def merge_font_props(default_props, original_props):
    """
    合并字体属性：优先使用原始文档中的字体属性，缺失时回退到默认值。

    参数：
        default_props: 默认字体属性 dict
        original_props: 从原始run提取的字体属性 dict

    返回：
        dict: 合并后的字体属性
    """
    merged = {}
    for key in ('ascii', 'hAnsi', 'eastAsia'):
        merged[key] = original_props.get(key) or default_props.get(key)
    merged['sz'] = original_props.get('sz') or default_props.get('sz')
    return merged


def apply_font_props_to_rpr(rPr, font_props, italic=None):
    """
    将字体属性写入 w:rPr 元素。

    参数：
        rPr: w:rPr lxml Element
        font_props: 字体属性 dict
        italic: 斜体标记，True=斜体, False=正体, None=不修改
    """
    if not font_props:
        font_props = {}

    rFonts = etree.SubElement(rPr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rFonts')
    if font_props.get('ascii'):
        rFonts.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}ascii', font_props['ascii'])
    if font_props.get('hAnsi'):
        rFonts.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}hAnsi', font_props['hAnsi'])
    if font_props.get('eastAsia'):
        rFonts.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}eastAsia', font_props['eastAsia'])

    if font_props.get('sz'):
        sz_elem = etree.SubElement(rPr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sz')
        sz_elem.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', font_props['sz'])
        szCs = etree.SubElement(rPr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}szCs')
        szCs.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', font_props['sz'])

    # 斜体处理
    if italic is not None:
        i_elem = etree.SubElement(rPr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}i')
        if not italic:
            i_elem.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', 'false')


def make_del_element(text, author=REVISION_AUTHOR, rev_id=1, font_props=None, italic=None):
    """
    创建 w:del 元素（修订模式的删除标记）

    参数：
        text: 要标记为删除的文本
        author: 修订者名称
        rev_id: 修订ID
        font_props: 字体属性 dict
        italic: 斜体标记，True=斜体, False=正体, None=保持原始（不作修改）

    返回：
        lxml Element: 删除修订元素
    """
    del_elem = etree.SubElement(
        etree.Element('root'),
        '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}del'
    )
    del_elem.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id', str(rev_id))
    del_elem.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}author', author)
    del_elem.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}date', REVISION_DATE)

    # 添加 run
    r = etree.SubElement(del_elem, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r')
    rPr = etree.SubElement(r, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rPr')
    apply_font_props_to_rpr(rPr, font_props, italic=italic)
    etree.SubElement(rPr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}delText')
    del_text = etree.SubElement(r, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}delText')
    del_text.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    del_text.text = text

    return del_elem


def make_ins_element(text, author=REVISION_AUTHOR, rev_id=1, font_props=None, italic=None):
    """
    创建 w:ins 元素（修订模式的插入标记）

    参数：
        text: 要标记为插入的文本
        author: 修订者名称
        rev_id: 修订ID
        font_props: 字体属性 dict
        italic: 斜体标记，True=斜体, False=正体, None=保持原始（不作修改）

    返回：
        lxml Element: 插入修订元素
    """
    ins_elem = etree.SubElement(
        etree.Element('root'),
        '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}ins'
    )
    ins_elem.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id', str(rev_id))
    ins_elem.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}author', author)
    ins_elem.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}date', REVISION_DATE)

    r = etree.SubElement(ins_elem, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r')
    rPr = etree.SubElement(r, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rPr')
    apply_font_props_to_rpr(rPr, font_props, italic=italic)
    t = etree.SubElement(r, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t')
    t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    t.text = text

    return ins_elem


def make_blue_comment_element(text, author=REVISION_AUTHOR, comment_id=1):
    """
    创建蓝色批注元素

    参数：
        text: 批注文本
        author: 批注作者
        comment_id: 批注ID

    返回：
        lxml Element: 批注元素
    """
    comment = etree.SubElement(
        etree.Element('root'),
        '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}comment'
    )
    comment.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id', str(comment_id))
    comment.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}author', author)
    comment.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}date', REVISION_DATE)
    comment.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}initials', 'YZZK')

    # 添加段落
    p = etree.SubElement(comment, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p')
    pPr = etree.SubElement(p, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pPr')
    r = etree.SubElement(p, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r')
    rPr = etree.SubElement(r, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rPr')

    # 设置蓝色 (0000FF)
    color = etree.SubElement(rPr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}color')
    color.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', '0000FF')

    t = etree.SubElement(r, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t')
    t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    t.text = text

    return comment


def find_text_in_footnote_run(footnote_elem, search_text):
    """
    在脚注 XML 元素中查找指定文本的位置。

    返回：
        (run_element, t_element, start_offset) 或 None
    """
    for run in footnote_elem.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r'):
        for t in run.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t'):
            if t.text and search_text in t.text:
                return (run, t, t.text.index(search_text))
    return None


def find_para_and_index(run_elem):
    """
    从 w:r 元素向上查找包含它的段落和它（或其包装元素）在段落中的位置。

    处理 run 嵌套在 <w:hyperlink>、<w:sdt> 等元素中的情况。
    返回最外层可移除的块级元素及其位置。

    参数：
        run_elem: lxml Element (w:r)

    返回：
        (block_to_remove, paragraph, index_in_para) 或 (None, None, -1)
    """
    WML = WML_NS
    skip_tags = {
        f'{{{WML}}}bookmarkStart',
        f'{{{WML}}}bookmarkEnd',
        f'{{{WML}}}commentRangeStart',
        f'{{{WML}}}commentRangeEnd',
        f'{{{WML}}}commentReference',
        f'{{{WML}}}proofErr',
    }
    block = run_elem
    while block is not None:
        parent = block.getparent()
        if parent is None:
            return None, None, -1
        if parent.tag == f'{{{WML}}}p':
            return block, parent, list(parent).index(block)
        # 跳过标记类元素（bookmark、commentRange等），继续向上查找
        if parent.tag in skip_tags:
            block = parent
            continue
        # hyperlink、sdt 等包装元素：作为块级移除
        block = parent
    return None, None, -1


def _set_deepcopy_run_text(run_elem, text):
    """
    设置 deepcopy 后的 run 的文本内容。

    处理一个 run 内可能包含多个 <w:t> 元素的情况：
    将文本放在第一个 <w:t>，删除其余 <w:t> 的文本内容。

    参数：
        run_elem: lxml Element (w:r, deepcopy)
        text: 要设置的文本
    """
    WML = WML_NS
    t_elements = run_elem.findall('.//{{{}}}t'.format(WML))
    for i, t in enumerate(t_elements):
        t.text = text if i == 0 else None


def apply_simple_text_replacement(footnote_elem, old_text, new_text, rev_id, italic=None):
    """
    执行简单的文本替换（修订模式）。

    在脚注 XML 中：
    1. 找到包含 old_text 的 w:t 元素
    2. 用 w:del（旧文本）+ w:ins（新文本）替换

    参数：
        footnote_elem: 脚注 lxml Element
        old_text: 要替换的旧文本
        new_text: 替换后的新文本
        rev_id: 修订ID
        italic: 斜体标记，True=斜体, False=正体, None=保持原始不作修改

    返回：
        bool: 是否成功找到并替换
    """
    result = find_text_in_footnote_run(footnote_elem, old_text)
    if not result:
        # 单 run 搜索失败，尝试跨 run 搜索
        if apply_cross_run_replacement(footnote_elem, old_text, new_text, rev_id, italic=italic):
            return True
        print(f"  警告：在脚注中未找到文本: '{old_text}'")
        return False

    run, t_elem, offset = result
    parent = t_elem.getparent()  # w:r（可能在 hyperlink/sdt 内）

    if parent is None:
        return False

    # 获取段落和在段落中的位置（处理嵌套 run 的情况）
    block_to_remove, para, block_index = find_para_and_index(parent)
    if para is None or block_index < 0:
        return False

    # 提取原始字体属性
    original_font = extract_font_props_from_run(parent)

    # 使用 run 的完整文本（上下文）判定语言，而非仅依赖 new_text
    # 这样中文上下文中的替换（如 and→&）也能获得中文默认字体
    context_text = t_elem.text if t_elem.text else old_text
    context_lang = detect_text_language(context_text)
    default_font = get_default_font_props(context_lang)
    font_props = merge_font_props(default_font, original_font)

    # 确定 del 元素的斜体（使用原始 run 的斜体状态）
    del_italic = original_font.get('italic') if italic is not None else None

    # 替换整个 run 内容的情况（简单情况：run 内容恰好是 old_text）
    if t_elem.text == old_text:
        del_elem = make_del_element(old_text, rev_id=rev_id, font_props=font_props, italic=del_italic)
        ins_elem = make_ins_element(new_text, rev_id=rev_id, font_props=font_props, italic=italic)

        para.remove(block_to_remove)
        para.insert(block_index, ins_elem)
        para.insert(block_index, del_elem)
        return True

    # 复杂情况：run 内容包含 old_text 外的文本
    full_text = t_elem.text
    before = full_text[:offset]
    after = full_text[offset + len(old_text):]

    # BUGFIX: 必须在 remove 之前保存位置
    para.remove(block_to_remove)
    insert_pos = block_index

    # before 部分
    if before:
        before_run = copy.deepcopy(parent)
        _set_deepcopy_run_text(before_run, before)
        para.insert(insert_pos, before_run)
        insert_pos += 1

    # del 部分
    del_elem = make_del_element(old_text, rev_id=rev_id, font_props=font_props, italic=del_italic)
    para.insert(insert_pos, del_elem)
    insert_pos += 1

    # ins 部分
    ins_elem = make_ins_element(new_text, rev_id=rev_id, font_props=font_props, italic=italic)
    para.insert(insert_pos, ins_elem)
    insert_pos += 1

    # after 部分
    if after:
        after_run = copy.deepcopy(parent)
        _set_deepcopy_run_text(after_run, after)
        para.insert(insert_pos, after_run)

    return True


def apply_cross_run_replacement(footnote_elem, old_text, new_text, rev_id, italic=None):
    """
    跨 run 文本替换（当单 run 搜索失败时的回退方案）。

    Word 文档中文字可能因格式差异被拆成多个 <w:r> 元素，
    此函数跨 run 拼接文本后匹配并替换。

    支持嵌套在 <w:hyperlink>、<w:sdt> 等元素中的 run。

    参数：
        footnote_elem: 脚注 lxml Element
        old_text: 要替换的旧文本
        new_text: 替换后的新文本
        rev_id: 修订ID
        italic: 斜体标记，True=斜体, False=正体, None=保持原始不作修改

    返回：
        bool: 是否成功找到并替换
    """
    WML = WML_NS
    for para in footnote_elem.findall('{{{wml}}}p'.format(wml=WML)):
        # 收集所有 run 及其文本（包括嵌套在 hyperlink/sdt 中的）
        # 每条: (containing_block, run_elem, run_text, global_start)
        segments = []
        char_pos = 0
        full_text = ""

        para_children = list(para)
        for child in para_children:
            tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
            if tag == 'r':
                # 直接的 run
                run_text = _extract_run_text(child, WML)
                if run_text:
                    segments.append((child, child, run_text, char_pos))
                    full_text += run_text
                    char_pos += len(run_text)
            elif tag in ('hyperlink', 'sdt'):
                # 嵌套的 run：从包装元素中提取
                for run in child.findall('.//{{{}}}r'.format(WML)):
                    run_text = _extract_run_text(run, WML)
                    if run_text:
                        segments.append((child, run, run_text, char_pos))
                        full_text += run_text
                        char_pos += len(run_text)

        match_idx = full_text.find(old_text)
        if match_idx == -1:
            continue

        match_end = match_idx + len(old_text)

        # 确定哪些 segment 与匹配区间重叠
        first_overlap = -1
        last_overlap = -1
        for si in range(len(segments)):
            seg_start = segments[si][3]
            seg_end = seg_start + len(segments[si][2])
            if seg_start < match_end and seg_end > match_idx:
                if first_overlap == -1:
                    first_overlap = si
                last_overlap = si

        if first_overlap == -1:
            continue

        # 提取字体属性，使用完整段落文本（上下文）判定语言
        original_font = extract_font_props_from_run(segments[first_overlap][1])
        context_lang = detect_text_language(full_text) if full_text else detect_text_language(old_text)
        default_font = get_default_font_props(context_lang)
        font_props = merge_font_props(default_font, original_font)

        # 确定 del 元素的斜体（使用原始 run 的斜体状态）
        del_italic = original_font.get('italic') if italic is not None else None

        # 收集要移除的块（每个 block 只移除一次）
        blocks_to_remove = {}  # para_index -> block_element
        for si in range(first_overlap, last_overlap + 1):
            block = segments[si][0]
            bidx = list(para).index(block)
            blocks_to_remove[bidx] = block

        # 记录第一个块的 para 位置作为插入起点
        first_block_idx = min(blocks_to_remove.keys())

        # 从后向前移除所有重叠块
        for bidx in sorted(blocks_to_remove.keys(), reverse=True):
            para.remove(blocks_to_remove[bidx])

        # 从前向后重建
        insert_pos = first_block_idx

        for si in range(first_overlap, last_overlap + 1):
            _, run_elem, run_text = segments[si][:3]
            global_start = segments[si][3]
            global_end = global_start + len(run_text)

            match_start_in_run = max(0, match_idx - global_start)
            match_end_in_run = min(len(run_text), match_end - global_start)

            before = run_text[:match_start_in_run]
            matched = run_text[match_start_in_run:match_end_in_run]
            after = run_text[match_end_in_run:]

            if before:
                before_run = copy.deepcopy(run_elem)
                _set_deepcopy_run_text(before_run, before)
                para.insert(insert_pos, before_run)
                insert_pos += 1

            if matched:
                del_elem = make_del_element(matched, rev_id=rev_id, font_props=font_props, italic=del_italic)
                para.insert(insert_pos, del_elem)
                insert_pos += 1

            if after:
                after_run = copy.deepcopy(run_elem)
                _set_deepcopy_run_text(after_run, after)
                para.insert(insert_pos, after_run)
                insert_pos += 1

        ins_elem = make_ins_element(new_text, rev_id=rev_id, font_props=font_props, italic=italic)
        para.insert(insert_pos, ins_elem)

        return True

    return False


def _extract_run_text(run_elem, WML):
    """从 w:r 元素中提取所有 <w:t> 文本（支持多 t 元素）。"""
    text = ''
    for t in run_elem.findall('.//{{{}}}t'.format(WML)):
        if t.text:
            text += t.text
    return text


def add_blue_comment_to_footnote_range(footnote_elem, comment_text, comment_id):
    """
    在脚注范围的起始和结束位置添加评论标记。
    将评论添加到 comments.xml（需在外部处理）。

    此函数在脚注段落中添加 w:commentRangeStart 和 w:commentRangeEnd 标记。

    参数：
        footnote_elem: 脚注 lxml Element
        comment_text: 批注文本
        comment_id: 批注ID

    返回：
        (comment_xml_element, comment_range_start, comment_range_end): 批注元素
    """
    # 获取脚注的第一个段落
    first_para = footnote_elem.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p')
    if first_para is None:
        return None

    # 创建 commentRangeStart
    start = etree.SubElement(first_para, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}commentRangeStart')
    start.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id', str(comment_id))

    # 将 commentRangeStart 移到段落开头
    first_para.insert(0, start)

    # 创建 commentRangeEnd（必须在 commentReference 之前，符合 OOXML 规范）
    end = etree.SubElement(first_para, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}commentRangeEnd')
    end.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id', str(comment_id))

    # 创建 commentReference（OOXML 规范要求必须放在 w:r 元素内，不能直接作为 w:p 的子元素）
    ref_run = etree.SubElement(first_para, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r')
    ref_run_rPr = etree.SubElement(ref_run, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rPr')
    ref = etree.SubElement(ref_run, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}commentReference')
    ref.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id', str(comment_id))

    # 创建批注元素
    comment_elem = make_blue_comment_element(comment_text, comment_id=comment_id)

    return comment_elem


def process_corrections(docx_path, corrections, output_path):
    """
    处理所有修订，生成新的 docx 文件。

    参数：
        docx_path: 原始 docx 文件路径
        corrections: 修订列表
        output_path: 输出文件路径

    修订格式：
    [
        {
            "footnote_id": "2",
            "rule": "第25条",
            "error_type": "缺少冒号",
            "old_text": "王名扬",
            "new_text": "王名扬：",
            "reason": "作者与书名之间应用冒号"
        }
    ]
    """
    # 复制原文件
    shutil.copy2(docx_path, output_path)

    rev_id = REVISION_ID_BASE
    comment_id = 1
    all_comments = []

    with zipfile.ZipFile(output_path, 'r') as z:
        footnotes_xml = z.read('word/footnotes.xml')
        # 读取或创建 comments.xml
        try:
            comments_xml = z.read('word/comments.xml')
            comments_tree = etree.fromstring(comments_xml)
        except KeyError:
            # 创建新的 comments.xml
            comments_root = etree.Element(
                '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}comments',
                nsmap={'w': WML_NS}
            )
            comments_tree = comments_root

        all_parts = {}
        for name in z.namelist():
            all_parts[name] = z.read(name)

    # 解析脚注 XML
    footnotes_tree = etree.fromstring(footnotes_xml)

    # 按脚注 ID 分组处理修订
    corrections_by_fn = {}
    for c in corrections:
        fn_id = c.get('footnote_id')
        if fn_id not in corrections_by_fn:
            corrections_by_fn[fn_id] = []
        corrections_by_fn[fn_id].append(c)

    for fn_id, fn_corrections in corrections_by_fn.items():
        # 找到对应脚注元素
        fn_found = False
        for fn_elem in footnotes_tree.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}footnote'):
            if fn_elem.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id') == str(fn_id):
                fn_found = True

                # 处理该脚注的每个修订
                for corr in fn_corrections:
                    old_text = corr.get('old_text', '')
                    new_text = corr.get('new_text', '')
                    rule = corr.get('rule', '')
                    reason = corr.get('reason', '')
                    format_changes = corr.get('format_changes', {})

                    # 解析斜体设置
                    italic = None
                    if isinstance(format_changes, dict) and 'italic' in format_changes:
                        italic = format_changes['italic']  # True=加斜体, False=去斜体(正体)

                    if corr.get('comment_only', False):
                        # 仅批注，不修改文本
                        success = False
                    else:
                        # 执行文本替换
                        success = apply_simple_text_replacement(fn_elem, old_text, new_text, rev_id, italic=italic)

                    # 添加蓝色批注
                    if success or corr.get('comment_only', False):
                        comment_text = f"[{rule}] {corr.get('error_type', '格式错误')}：{reason}"
                        if old_text and new_text and old_text != new_text:
                            comment_text += f"\n修正：\"{old_text}\" → \"{new_text}\""
                        if isinstance(format_changes, dict) and format_changes:
                            fmt_desc = []
                            fc_italic = format_changes.get('italic')
                            if fc_italic is True:
                                fmt_desc.append('应用斜体')
                            elif fc_italic is False:
                                fmt_desc.append('改用正体')
                            if fmt_desc:
                                comment_text += f"\n格式修正：{'，'.join(fmt_desc)}"

                        comment_elem = add_blue_comment_to_footnote_range(
                            fn_elem, comment_text, comment_id
                        )

                        if comment_elem is not None:
                            all_comments.append(comment_elem)    # comment_elem 本身就是 <w:comment> 元素
                            comment_id += 1

                    rev_id += 1
                break

        if not fn_found:
            print(f"  警告：未找到脚注 ID={fn_id}")

    # 更新 comments.xml
    comments_tree = etree.Element(
        '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}comments',
        nsmap={'w': WML_NS}
    )
    for cmt in all_comments:
        comments_tree.append(cmt)

    # 重新打包 docx
    all_parts['word/footnotes.xml'] = etree.tostring(
        footnotes_tree, xml_declaration=True, encoding='UTF-8', standalone=True
    )
    all_parts['word/comments.xml'] = etree.tostring(
        comments_tree, xml_declaration=True, encoding='UTF-8', standalone=True
    )

    # 更新 [Content_Types].xml 以确保包含 comments
    content_types_xml = all_parts.get('[Content_Types].xml', b'')
    if content_types_xml:
        ct_tree = etree.fromstring(content_types_xml)
        # 检查是否已有 comments 的内容类型
        has_comments_ct = False
        for override in ct_tree.iter('{http://schemas.openxmlformats.org/package/2006/content-types}Override'):
            if override.get('PartName') == '/word/comments.xml':
                has_comments_ct = True
                break
        if not has_comments_ct:
            override = etree.SubElement(ct_tree, '{http://schemas.openxmlformats.org/package/2006/content-types}Override')
            override.set('PartName', '/word/comments.xml')
            override.set('ContentType', 'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml')
        all_parts['[Content_Types].xml'] = etree.tostring(
            ct_tree, xml_declaration=True, encoding='UTF-8', standalone=True
        )

    # 写入新 docx
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zout:
        for name, data in all_parts.items():
            zout.writestr(name, data)

    print(f"\n修订完成。共处理 {len(corrections)} 处修订，生成 {comment_id - 1} 条批注。")
    print(f"输出文件: {output_path}")


def main():
    parser = argparse.ArgumentParser(description='对 .docx 文件中的脚注执行修订模式修改和蓝色批注')
    parser.add_argument('input', type=str, help='输入的 .docx 文件路径')
    parser.add_argument('--corrections', '-c', type=str, required=True, help='修订信息 JSON 文件路径')
    parser.add_argument('--output', '-o', type=str, required=True, help='输出的 .docx 文件路径')

    args = parser.parse_args()

    with open(args.corrections, 'r', encoding='utf-8') as f:
        corrections = json.load(f)

    process_corrections(args.input, corrections, args.output)


if __name__ == '__main__':
    main()

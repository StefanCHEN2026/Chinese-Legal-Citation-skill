#!/usr/bin/env python3
"""
法学论文引注检查工具模块
提供引注语言判定、类型识别、格式检查等辅助函数，供核查时使用。

依赖：无（纯标准库）
"""

import re
import json
from typing import List, Tuple, Optional, Dict


# ── CJK 字符范围 ──
CJK_RE = re.compile(r'[\u4e00-\u9fff\u3400-\u4dbf\uf900-\ufaff]')

# ── 中文标点 ──
CHINESE_PUNCTUATION = re.compile(r'[，。；：？！""''《》（）【】—…、·]')

# ── 英文标点 ──
ENGLISH_PUNCTUATION = re.compile(r'[,.;:?!""''()\[\]{}…]')

# ── 各国文字特征 ──
JAPANESE_KANA_RE = re.compile(r'[\u3040-\u309f\u30a0-\u30ff]')  # 假名
FRENCH_FEATURES = re.compile(r'\b(éd\.|édition|éditions)\b')
GERMAN_FEATURES = re.compile(r'\b(Aufl\.|Verlag|Rn\.)\b')


def is_chinese_citation(text: str) -> bool:
    """判断引注是否包含中文"""
    return bool(CJK_RE.search(text))


def is_english_citation(text: str) -> bool:
    """
    判断引注是否为英文。
    不含中文字符、不含假名、不含法文/德文特征词。
    """
    if CJK_RE.search(text):
        return False
    if JAPANESE_KANA_RE.search(text):
        return False
    if FRENCH_FEATURES.search(text):
        return False
    if GERMAN_FEATURES.search(text):
        return False
    # 以英文字母开头或包含典型英文出版社/期刊名
    if re.match(r'^[A-Za-z]', text.strip()):
        return True
    return False


# ── 中文引注类型判定 ──

def classify_chinese_citation(text: str) -> str:
    """
    将中文引注文本分类。

    返回类型：
        - 'internet'      网络文献
        - 'case'          司法案例
        - 'legal'         法律文件
        - 'ancient'       古籍
        - 'article_journal' 期刊文章
        - 'article_newspaper' 报纸文章
        - 'article_collection' 文集文章
        - 'dissertation'  学位论文
        - 'book'          图书
        - 'repeated'      再次引用
        - 'unknown'       无法判定
    """
    text = text.strip()

    # 1. 网络文献 (优先级最高)
    if re.search(r'https?://', text):
        return 'internet'
    if '微信公众号' in text or '载微信公众号' in text:
        return 'internet'
    if '载' in text and ('博客' in text or 'Blog' in text):
        return 'internet'
    if '载' in text and 'App' in text:
        return 'internet'
    if '访问' in text and ('年' in text and '月' in text and '日' in text):
        if re.search(r'https?://', text):
            return 'internet'

    # 2. 再次引用
    if text.startswith('同前注') or text.startswith('同上注') or text == '同上':
        return 'repeated'
    if text.startswith('前引') and ('注' in text[:10]):
        return 'repeated'

    # 3. 司法案例
    case_patterns = [
        r'案[),，.]',  # "xxx案"
        r'诉.*案[),，.]',  # "原告诉被告xxx案"
        r'人民法院.*判决书',
        r'人民法院.*裁定书',
        r'人民法院.*决定书',
        r'[\（(]\d{4}[）)]\S*号',
        r'指导案例\s*\d+号',
        r'公报.*年第\d+期',
        r'人民法院案例库',
        r'刑事判决|民事判决|行政判决',
        r'行政复议决定',
        r'行政处罚决定',
    ]
    for pat in case_patterns:
        if re.search(pat, text):
            return 'case'

    # 4. 法律文件
    legal_patterns = [
        r'《[^》]+法》',
        r'《[^》]+条例》',
        r'《[^》]+办法》',
        r'《[^》]+规定》',
        r'《[^》]+决定》',
        r'《[^》]+通知》',
        r'《[^》]+意见》',
        r'法释[\〔\[]',
        r'国发[\〔\[]',
        r'国办发[\〔\[]',
        r'^\S+[部委局院署会]《',
        r'^\S+[部委局院署会]发布',
        r'联合国.*公约|条约|宣言',
    ]
    for pat in legal_patterns:
        if re.search(pat, text):
            return 'legal'

    # 5. 古籍
    ancient_patterns = [
        r'^[（(][唐宋元明清秦汉魏晋隋][）)]',
        r'^[（(]春秋[）)]|^[（(]战国[）)]',
        r'《[^》]+[注疏集解训诂考]》',
        r'^《[^》]+》$',  # 纯古籍名如《论语·述而》
    ]
    for pat in ancient_patterns:
        if re.search(pat, text):
            return 'ancient'

    # 如果出现"卷"+"刻本"或"卷"+"影印"或"卷"+"活字"等，也是古籍
    if re.search(r'卷[^》]*(刻本|影印|活字|抄本)', text):
        return 'ancient'

    # 6. 学位论文
    if ('博士学位论文' in text or '硕士学位论文' in text or
        '博士论文' in text or '硕士论文' in text):
        return 'dissertation'

    # 7. 期刊/报纸/文集文章
    # 注意：含有"载"后面接期刊、报纸或文集名称
    if '载《' in text or '载' in text:
        # 报纸特征：有具体的日（日报）
        if re.search(r'年\d{1,2}月\d{1,2}日', text):
            return 'article_newspaper'
        # 文集特征："载""主编""编""文集""论文集"
        if re.search(r'载.*(主编|编|编译|文集|论文集|论丛)', text):
            return 'article_collection'
        # 期刊特征："年第"或"卷"或"期"
        if re.search(r'\d{4}年第\d+期', text) or re.search(r'第\d+卷第\d+期', text):
            return 'article_journal'
        # 默认期刊
        if re.search(r'第\d+[卷辑]', text):
            return 'article_collection'
        return 'article_journal'

    # 8. 图书
    book_patterns = [
        r'出版社',
        r'年版\s*[，,.]',
        r'年版$',
        r'第\d+版',
    ]
    for pat in book_patterns:
        if re.search(pat, text):
            return 'book'

    # 9. 翻译作品（在图书基础上）
    # 已经在 book 类型中涵盖

    return 'unknown'


# ── 英文引注类型判定 ──

def classify_english_citation(text: str) -> str:
    """
    将英文引注文本分类。

    返回类型：
        - 'internet'      网络文献
        - 'case'          案例
        - 'article_journal' 期刊文章
        - 'book'          图书
        - 'repeated'      再次引用
        - 'unknown'       无法判定
    """
    text = text.strip()

    # 1. 网络文献
    if re.search(r'https?://', text):
        return 'internet'
    if 'visited on' in text.lower():
        return 'internet'

    # 2. 再次引用
    if text.startswith('Ibid') or text.startswith('ibid'):
        return 'repeated'
    if re.search(r'supra\s+note\s+\d+', text, re.IGNORECASE):
        return 'repeated'

    # 3. 案例
    case_patterns = [
        r'\bv\.\b',           # 原告诉被告
        r'\bU\.S\.\s+\d+',    # 美国案例报告
        r'\bF\.\d+d\s+\d+',   # 联邦案例报告
        r'\bICJ\b',           # 国际法院
        r'\bICSID\b',         # ICSID
        r'\bAppellate\s+Body\b',  # WTO 上诉机构
        r'\bEWCA\b|\bEWHC\b|\bUKSC\b|\bUKHL\b',  # 英国法院
        r'\[?\d{4}\]?\s+\w+\s+\d+',  # 案例报告格式
    ]
    for pat in case_patterns:
        if re.search(pat, text):
            return 'case'

    # 4. 期刊文章
    # 特征：期刊名 + 卷号 + 页码
    journal_patterns = [
        r'\d+\s+\w[\w\s]+\w+\s+\d+',  # 卷次 期刊名 页码
        r'Vol\.\s*\d+',               # Vol. XX
        r'\bJournal\b|\bReview\b|\bLaw\b.*\d+',  # 期刊特征词
    ]
    for pat in journal_patterns:
        if re.search(pat, text):
            return 'article_journal'

    # 5. 图书
    book_patterns = [
        r'\bPress\b',
        r'\bPublishing\b',
        r'\bPublisher\b',
        r'\bInc\.\b',
        r',\s*\d{4}[,.]',
        r'\(\d+th\s+ed\.\)',
        r'\bed\.?\s*,',
    ]
    for pat in book_patterns:
        if re.search(pat, text):
            return 'book'

    return 'unknown'


# ── 格式检查辅助函数 ──

def check_author_title_separator(text: str) -> Optional[str]:
    """
    检查中文引注中作者与书名之间是否有冒号。
    返回 None 表示正确，否则返回错误描述。
    """
    # 匹配模式：作者名 + 书名
    # 正确格式：王名扬：《美国行政法》
    # 错误格式：王名扬《美国行政法》、王名扬"美国行政法"
    m = re.match(r'^[\u4e00-\u9fff·\w\s]+[：:]', text)
    if not m:
        return "作者与书名之间缺少冒号"
    return None


def check_book_title_format(text: str) -> List[str]:
    """
    检查中文书名号使用是否正确。
    """
    errors = []
    # 书名是否使用了《》而非 "" 或 <>
    if re.search(r'"[\u4e00-\u9fff]+"', text):
        errors.append("书名使用了英文引号，应使用书名号《》")
    return errors


def check_page_number_format(text: str) -> Optional[str]:
    """
    检查中文页码格式是否为"第X页"格式。
    """
    # 含有页码但格式不对
    if re.search(r'[\d]+页', text) and not re.search(r'第[\d]+页', text):
        return '页码缺少"第"字，正确格式为"第X页"'
    if re.search(r'[\d]+-\d+页', text) and not re.search(r'第[\d]+-[\d]+页', text):
        return '起止页码缺少"第"字'
    return None


def check_zai_character(text: str) -> Optional[str]:
    """
    检查引用文章时是否有"载"字。
    适用于期刊文章、文集文章。
    """
    # 如果是文章类型但没有"载"字
    if re.search(r'《[^》]+》(?!载)[，,\s]*《[^》]+》\d{4}', text):
        return '引注文章缺少"载"字（应在来源期刊/文集前加"载"）'
    return None


def check_publisher_year_format(text: str) -> Optional[str]:
    """
    检查出版社+年份格式是否正确。
    正确："出版社1995年版"，错误："出版社1995年"、"出版社1995"
    """
    if re.search(r'出版社\s*\d{4}[^年]', text) and '年版' not in text:
        return '出版年份缺少"年版"，正确格式为"出版社1995年版"'
    # 检查"1995年版，第X页"格式
    if re.search(r'出版社\d{4}\s*[，,]\s*第', text):
        return '出版社与年份之间缺少连接，应为"出版社1995年版"而非"出版社，1995"'
    return None


def check_translator_format(text: str) -> List[str]:
    """
    检查翻译作品的格式。
    检查项：国籍标记、译者标注。
    """
    errors = []
    # 检查是否有外文作者特征但没有国籍标记
    # 有外国特征姓名但没有 [国籍] 标记
    foreign_name_pattern = re.search(r'[（(]?\s*([A-Z][a-z]+(?:\s+[A-Z][a-z]+)+)\s*[）)]?', text)
    if foreign_name_pattern and not re.search(r'\[[^\]]+\]', text):
        # 可能是翻译作品但缺国籍标记
        if '译' in text:
            errors.append('翻译作品缺少国籍标记，如[美]、[德]')

    # 检查译者是否存在但格式不完整
    if re.search(r'译[者]?\s*[，,]', text) or '译' in text:
        if not re.search(r'，\S+译[，,\s]', text) and not re.search(r'，\S+译$', text):
            if not re.search(r'译，|译。|译$', text):
                errors.append('译者信息格式可能有误，应为"译者译"')

    return errors


def check_chinese_english_punctuation_mix(text: str) -> List[str]:
    """
    检查中英文标点混用问题。
    """
    errors = []

    # 中文文本中不应出现英文逗号（除非是URL）
    if is_chinese_citation(text) and not re.search(r'https?://', text):
        # 中文逗号 vs 英文逗号
        if ',' in text.split('http')[0] if 'http' in text else ',' in text:
            errors.append('中文引注中使用了英文逗号，应使用中文逗号（，）')

    # 英文字母上下文检查混用
    if re.search(r'[\u4e00-\u9fff]\s*,\s*[\u4e00-\u9fff]', text):
        errors.append('中文字符间使用了英文逗号')

    return errors


def check_bracket_format(text: str) -> Optional[str]:
    """
    检查括号类型是否正确。
    六角括号〔〕用于文号年份，圆括号（）用于普通内容，
    方括号[]用于国籍标记。
    """
    # 文号年份应该用六角括号
    if re.search(r'[（(]\d{4}[）)]\d+号', text):
        return '文号年份应使用六角括号〔〕，而非圆括号'
    if re.search(r'\[\d{4}\]\d+号', text):
        return '文号年份应使用六角括号〔〕，而非方括号'
    return None


def check_author_nationality_bracket(text: str) -> Optional[str]:
    """
    检查国籍标记是否使用方括号 []。
    正确：[美]、[德]、[日]
    错误：（美）、【美】、[美]
    """
    # 检查是否有国籍标记但用了错误的括号
    if re.search(r'[（(]美[）)]|[（(]德[）)]|[（(]日[）)]|[（(]英[）)]|[（(]法[）)]', text):
        return '国籍标记应使用方括号如[美]，而非圆括号（美）'
    if re.search(r'【美】|【德】|【日】|【英】|【法】', text):
        return '国籍标记应使用方括号如[美]，而非方头括号【美】'
    return None


def extract_citation_components(text: str) -> Dict[str, str]:
    """
    提取中文引注的各个组成部分。
    返回 components dict，key 包括: author, title, translator, source, publisher, year, page
    """
    components = {
        'author': '',
        'title': '',
        'translator': '',
        'source': '',
        'publisher': '',
        'year': '',
        'page': '',
        'raw': text
    }

    # 提取作者（冒号之前的部分）
    author_match = re.match(r'^([^：:]+)[：:]', text)
    if author_match:
        components['author'] = author_match.group(1).strip()

    # 提取书名（第一个书名号中的内容）
    title_match = re.search(r'《([^》]+)》', text)
    if title_match:
        components['title'] = title_match.group(1)

    # 提取出版社
    pub_match = re.search(r'([^\s，,]+出版社)', text)
    if pub_match:
        components['publisher'] = pub_match.group(1)

    # 提取年份
    year_match = re.search(r'(\d{4})年', text)
    if year_match:
        components['year'] = year_match.group(1)

    # 提取页码
    page_match = re.search(r'第([\d]+(?:\s*[-—–]\s*[\d]+)?)页', text)
    if page_match:
        components['page'] = f"第{page_match.group(1)}页"

    # 提取译者
    trans_match = re.search(r'，([^，]+译)', text)
    if trans_match:
        components['translator'] = trans_match.group(1)

    return components


def generate_checklist_for_type(citation_type: str) -> List[str]:
    """
    根据引注类型返回应检查的项目清单（供 Claude 参考）。
    """
    checklist_map = {
        'book': [
            '作者与书名之间是否有冒号（：）',
            '书名是否使用书名号《》',
            '书名是否完整（含副标题？）',
            '是否有译者信息（翻译作品）',
            '译者是否标注"译"字',
            '国籍标记是否使用方括号[]',
            '出版社名称是否完整',
            '版本信息格式（"年版"或"第X版"）',
            '出版年份只写年份不写月份',
            '页码格式："第X页"或"第X-X页"',
            '各要素之间用逗号分隔',
            '结尾应有句号',
        ],
        'article_journal': [
            '文章标题是否使用书名号《》',
            '载《期刊名》前的"载"字是否有',
            '期刊名是否用书名号',
            '期刊版别名称是否在书名号内（如"（哲学社会科学版）"）',
            '出版信息是否为"年份第X期"格式',
            '页码格式："第X页"或"第X-X页"',
        ],
        'article_newspaper': [
            '文章标题是否使用书名号《》',
            '"载"字是否在报纸名前',
            '日期格式是否为"年X月X日"',
            '版面信息格式是否为"第X版"（不加"0"）',
        ],
        'article_collection': [
            '文章标题是否使用书名号《》',
            '"载"字是否在文集信息前',
            '文集编者信息是否完整（主编/编/编译）',
            '文集书名是否用书名号',
            '出版社和年份是否完整',
        ],
        'case': [
            '案例名称是否准确（原告诉被告案由案）',
            '案号格式（法院名+年份+案号+文书类型）',
            '案号年份是否用圆括号（）',
            '案例来源标注是否正确',
            '指导性案例格式（指导案例X号（年份））',
        ],
        'legal': [
            '法律名称是否用书名号《》',
            '"试行""草案"是否在书名号内',
            '"中华人民共和国"是否可省略',
            '条款序数是否为阿拉伯数字',
            '修改版本是否标注修改年份',
            '已失效文件是否标注"已废止"',
            '文号年份是否用六角括号〔〕',
            '规范性文件是否标明制定机关',
        ],
        'internet': [
            '是否有网站名称',
            '是否有上传日期',
            '链接前是否有逗号',
            '是否需要标注访问日期',
            '微信公众号格式：载微信公众号"名称"日期',
            '微博格式：微博账号，日期，链接',
        ],
        'dissertation': [
            '格式：作者：《题名》，学位授予单位 年份学位层级，第X页',
            '是否标明学位层级（博士/硕士）',
            '是否标明学位授予单位',
            '是否标明答辩年份',
        ],
        'ancient': [
            '作者朝代是否用圆括号（朝代）',
            '卷册是否写在书名后不加括号',
            '版本信息是否标明（刻本/影印本/整理本）',
            '出版信息是否完整',
        ],
        'repeated': [
            '"同前注[X]"格式是否正确',
            '作者名是否正确（不标国籍/朝代）',
            '文献名是否正确',
            '"同上注"/"Ibid."使用是否正确',
        ],
    }

    return checklist_map.get(citation_type, ['通用格式检查：标点符号、要素完整性'])


if __name__ == '__main__':
    # 简单测试
    test_cases = [
        ("王名扬：《美国行政法》，中国法制出版社1995年版，第18页。", "book"),
        ("张明楷：《刑法学的思路》，载《中国法学》2020年第3期，第45页。", "article_journal"),
        ("Roe v. Wade, 410 U.S. 113 (1973).", "case"),
        ("Charles A. Reich, The New Property, 73 Yale Law Journal 733 (1964).", "article_journal"),
        ('《民法典》第1224条第1款第2项', 'legal'),
        ('梁秋坪、郝萍：《全国打击治理农村赌博工作现场会召开》，载人民网2024年10月12日，http://example.com', 'internet'),
        ("同前注[16]，耶林：《为权利而斗争》，第3页。", "repeated"),
        ("同上注。", "repeated"),
    ]

    for text, expected in test_cases:
        if is_chinese_citation(text):
            result = classify_chinese_citation(text)
        else:
            result = classify_english_citation(text)
        status = "OK" if result == expected else f"FAIL (expected {expected})"
        print(f"[{status}] {text[:60]}... -> {result}")

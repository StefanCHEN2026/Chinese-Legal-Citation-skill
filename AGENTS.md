# 法学论文脚注引注格式核查与修订

## 概述

此项目用于核查法学论文 .docx 文件中脚注的引注格式是否符合《法学引注手册》规范，并以 Word 修订模式（Track Changes）进行格式修正，同时添加蓝色批注说明修改原因。

**只修改格式，不修改内容。**

## 支持的语言

- 中文引注（含各类中文文献）
- 英文引注

不支持日文、法文、德文等其他语言的引注核查。

## 文件结构

```
├── AGENTS.md                          # 本文件，工具无关的Skill定义
├── CLAUDE.md                          # Claude Code 专用 Skill 定义
├── LICENSE                            # MIT 许可证
├── README.md                          # GitHub 项目说明
├── CORRECTIONS_SCHEMA.md              # corrections.json 格式规范
├── references/                        # 引注规则库
│   ├── reference-index.md             # 快速索引（按类型、按错误反查）
│   ├── rules-general.md               # 通用规则（第4-24条）
│   ├── rules-chinese-books.md         # 中文图书（第25-49、64条）
│   ├── rules-chinese-articles.md      # 中文已发表文章（第42-49条）
│   ├── rules-chinese-internet.md      # 中文网络文献（第50-58条）
│   ├── rules-chinese-unpublished.md   # 中文未发表文献（第59-65条）
│   ├── rules-chinese-legal.md         # 中文法律文件（第66-86条）
│   ├── rules-chinese-cases.md         # 中文司法案例（第87-94条）
│   ├── rules-english.md               # 英文引注（第95-115条）
│   └── rules-repeated-citation.md     # 再次引用（第13、101条）
├── scripts/
│   ├── docx_parser.py                 # 脚注提取工具
│   ├── docx_revisor.py                # 修订模式修改工具
│   └── citation_checker.py            # 引注分类与检查辅助
└── CORRECTIONS_SCHEMA.md              # corrections.json 格式规范
```

## 工具依赖

```bash
pip install python-docx lxml
```

## 工作流程

当用户请求核查某篇法学论文的脚注时，按以下步骤操作：

### 第一步：提取脚注

```bash
python scripts/docx_parser.py <论文.docx> -o footnotes.json
```

### 第二步：逐条核查

读取 `footnotes.json`，对每条脚注按以下子步骤处理：

#### 2a. 判断是否为引注

- **引注**：包含文献出处信息（作者名、文献名、出版社、期刊名、链接等），需要核查格式
- **说明**：纯文字解释或补充说明，不需要核查格式，跳过

> 引注类脚注中也常混有说明性文字，需要先分割，只对引注部分核查。

#### 2b. 判定引注语言

- **中文引注**：包含任意 CJK 字符（U+4E00-9FFF）
- **英文引注**：不含 CJK 字符，以英文字母开头

其他语言仅提示"非中英文引注，暂不支持核查"，跳过。

#### 2c. 判定引注类型

**中文引注类型（按优先级判定）：**

1. **网络文献** (`internet`)：含 HTTP/HTTPS 链接、微信公众号、博客、App
2. **再次引用** (`repeated`)：以"同前注""同上注""同上"开头
3. **司法案例** (`case`)：含"诉""案"+案号/判决书，或"人民法院""指导案例"
4. **法律文件** (`legal`)：含法律文件名或文号（法释、国发等）
5. **古籍** (`ancient`)：作者前有（朝代）标记，或含"卷""刻本""影印"
6. **学位论文** (`dissertation`)：含"博士学位论文""硕士学位论文"
7. **报纸文章** (`article_newspaper`)：含"载《》"+年月日
8. **文集文章** (`article_collection`)：含"载"+主编/编/文集/论丛
9. **期刊文章** (`article_journal`)：含"载《》"+年份期次
10. **图书** (`book`)：含"出版社"+"年版"或"第X版"
11. **未知** (`unknown`)：无法判定，通用规则检查

**英文引注类型（按优先级判定）：**

1. **网络文献** (`internet`)：含 HTTP/HTTPS 链接、"visited on"
2. **再次引用** (`repeated`)：以"Ibid."开头，或含"supra note"
3. **司法案例** (`case`)：含"v."、U.S. + 数字、ICJ、ICSID、UKSC 等
4. **期刊文章** (`article_journal`)：含"Vol."、期刊特征词
5. **图书** (`book`)：含"Press""Publishing""ed."等

详见 `scripts/citation_checker.py`。

#### 2d. 查找对应规则

| 类型 | 规则文件 |
|------|---------|
| 通用规则 | `references/rules-general.md` |
| 中文-图书 | `references/rules-chinese-books.md` |
| 中文-期刊/报纸/文集 | `references/rules-chinese-articles.md` |
| 中文-网络文献 | `references/rules-chinese-internet.md` |
| 中文-未发表文献 | `references/rules-chinese-unpublished.md` |
| 中文-法律文件 | `references/rules-chinese-legal.md` |
| 中文-司法案例 | `references/rules-chinese-cases.md` |
| 英文-全部 | `references/rules-english.md` |
| 再次引用 | `references/rules-repeated-citation.md` |

#### 2e. 逐条核查格式

重点检查：

**中文常见错误：**
- 作者名后缺冒号
- 书名未用书名号《》
- 缺少"载"字
- 缺少出版年份
- 页码格式缺"第"和"页"
- 译作缺国籍或译者
- 国籍标记括号错误（应用方括号[]）
- 网络文献缺上传/访问日期
- 法律文件书名号误用
- 文号年份括号错误（应用六角括号〔〕）
- 案例案号/来源格式错误
- 中英文标点混用

**英文常见错误：**
- 英文书名/文章名未用斜体
- 英文期刊缺卷号
- 英文引注缺 See
- 英文页码 p. 写成 pp.
- 作者连接词 and 改为 &
- 英文日期格式误用

核查原则：
- 只报告格式错误，不质疑内容
- 手册未明确规定的，不强行要求
- 标注对应手册条款编号

### 第三步：生成修正数据

对每条错误生成修正记录，格式见 `CORRECTIONS_SCHEMA.md`：

```json
{
  "footnote_id": "3",
  "rule": "第96条",
  "error_type": "作者格式错误",
  "old_text": "Yan and Hyman",
  "new_text": "Yan & Hyman",
  "reason": "两位作者之间应使用 & 而非 and"
}
```

- `old_text` 必须是脚注中实际存在的文本片段
- 只替换格式相关的标点、连接词、标记符号

### 第四步：报告核查结果

向用户报告核查结果摘要：

```
## 核查结果摘要

- 脚注总数：32
- 引注脚注：28
- 说明性脚注：4（跳过）
- 发现错误：15 处
- 涉及脚注：10 条
```

### 第五步：备份并应用修订

**安全原则：不修改用户原始文件。修订工具始终在副本上工作。**

```bash
python scripts/docx_revisor.py <论文.docx> -c corrections.json -o <论文_已修订.docx>
```

`docx_revisor.py` 内部先 `shutil.copy2(输入, 输出)` 再修改，原文件不会被触及。

### 第六步：输出结果

告知用户输出文件路径、修订数量和批注数量。

## 注意事项

1. **不修改内容**：只修改标点符号、连接词、括号类型、格式标记
2. **解释性文字不核查**：脚注中说明性文字跳过
3. **多文献脚注**：逐个核查每条引注
4. **外文文献优先中译本**：按手册第95条优先引用中译本
5. **不确定时提示**：标注"建议人工复核"
6. **同一错误逐条列出**：不合并

## 跨工具兼容性

本项目支持以下 AI 编程工具：

| 工具 | 使用方式 | 配置文件 |
|------|---------|---------|
| **Claude Code** | 打开项目目录即自动加载 | `CLAUDE.md` |
| **Codex (OpenAI)** | 打开项目目录即自动加载 | `AGENTS.md` |
| **Workbuddy** | 打开项目目录即自动加载 | `AGENTS.md` |
| **Cursor** | 打开项目目录即自动加载 | `AGENTS.md` + `.cursorrules` |
| **Windsurf** | 打开项目目录即自动加载 | `AGENTS.md` |
| **GitHub Copilot** | 打开项目目录即自动加载 | `AGENTS.md` |

所有工具使用相同的工作流程，只需将你的论文 .docx 文件放在项目目录中，然后告诉 AI："请核查这篇论文的脚注：`你的论文.docx`"。

# 法学引注手册 Skill

基于《法学引注手册》（2025版）的 AI Skill，用于自动核查法学论文 .docx 文件中脚注引注格式的正确性，并以 **Word 修订模式**（Track Changes）修正格式错误，添加**蓝色批注**说明修改依据。

> 只修改格式，不修改内容。

## 支持的 AI 工具

本项目兼容多种 AI 编程工具，开箱即用：

| 工具 | 入口 |
|------|------|
| **Claude Code** | 打开目录即自动加载 [`CLAUDE.md`](CLAUDE.md) |
| **Codex (OpenAI)** | 打开目录即自动加载 [`AGENTS.md`](AGENTS.md) |
| **Workbuddy** | 打开目录即自动加载 [`AGENTS.md`](AGENTS.md) |
| **Cursor** | 打开目录即自动加载 [`AGENTS.md`](AGENTS.md) |
| **Windsurf** | 打开目录即自动加载 [`AGENTS.md`](AGENTS.md) |
| **GitHub Copilot** | 打开目录即自动加载 [`AGENTS.md`](AGENTS.md) |

## 功能

- 从 .docx 文件中提取所有脚注
- 自动判定引注语言（中文/英文）和类型（期刊/图书/案例/法律文件等 11+6 种）
- 对照《法学引注手册》规则库逐条核查格式规范
- 以 **Word 修订模式** 修正格式错误
- 添加 **蓝色批注** 说明修改原因和依据条款
- 支持中文宋体、英文 Times New Roman 字体

## 快速开始

### 1. 安装依赖

```bash
git clone https://github.com/StefanCHEN2026/-2025--skill.git
cd -2025--skill
pip install -r requirements.txt
```

### 2. 使用

将法学论文 .docx 文件放入项目目录，然后在任意支持的 AI 工具中：

```
请核查这篇论文的脚注：我的论文.docx
```

AI 将自动完成：
1. 提取全部脚注
2. 逐条判定类型和语言
3. 对照规则库核查格式
4. 生成修正清单供确认
5. 以修订模式写入修改

### 3. 手动使用命令行工具

```bash
# 提取脚注
python scripts/docx_parser.py 论文.docx -o footnotes.json

# 应用修正（需先准备好 corrections.json）
python scripts/docx_revisor.py 论文.docx -c corrections.json -o 论文_已修订.docx
```

## 支持的引注类型和规则

| 类型 | 规则条款 | 语言 |
|------|---------|------|
| 通用规则（标点、引领词、夹注等） | 第4-24条 | 中/英 |
| 中文图书 | 第25-49条 | 中文 |
| 中文期刊/报纸/文集文章 | 第42-49条 | 中文 |
| 中文网络文献（含公众号） | 第50-58条 | 中文 |
| 中文未发表文献（学位论文、档案） | 第59-65条 | 中文 |
| 中文法律文件 | 第66-86条 | 中文 |
| 中文司法案例 | 第87-94条 | 中文 |
| 英文引注（著作/期刊/法律/案例） | 第95-115条 | 英文 |
| 再次引用（同上注、同前注、supra） | 第13、101条 | 中/英 |

## 项目结构

```
├── CLAUDE.md                     # Claude Code Skill 定义
├── AGENTS.md                     # Codex/Workbuddy/Cursor 等工具 Skill 定义
├── README.md                     # 本文件
├── LICENSE                       # MIT 许可证
├── CORRECTIONS_SCHEMA.md         # corrections.json 格式规范
├── requirements.txt              # Python 依赖
├── references/                   # 引注规则库（10个文件）
│   ├── reference-index.md        # 快速索引（按类型、按错误反查）
│   ├── rules-general.md          # 通用规则
│   ├── rules-chinese-books.md    # 中文图书
│   ├── rules-chinese-articles.md # 中文期刊/报纸/文集
│   ├── rules-chinese-internet.md # 网络文献
│   ├── rules-chinese-unpublished.md # 未发表文献
│   ├── rules-chinese-legal.md    # 法律文件
│   ├── rules-chinese-cases.md    # 司法案例
│   ├── rules-english.md          # 英文引注
│   └── rules-repeated-citation.md # 再次引用
└── scripts/                      # Python 工具
    ├── docx_parser.py            # 脚注提取
    ├── docx_revisor.py           # 修订模式修改
    └── citation_checker.py       # 引注分类辅助
```

## 核查范围

### 会检查的格式问题

- 标点符号（中英文标点混用、冒号/逗号/分号）
- 作者格式（国籍标记括号、作者连接词）
- 文献名格式（书名号/引号/斜体）
- 出版信息（出版社/年份/版次格式）
- 页码格式（"第X页"格式、p./pp.）
- 期刊格式（"载"字、卷期号）
- 法律文件格式（条款序数、文号括号类型）
- 案例格式（案号、案例来源）
- 网络文献格式（上传日期、访问日期）
- 再次引用格式（同前注/同上注/supra/Ibid.）

### 不会检查

- 文献名称、作者姓名、出版社名称是否正确（内容核查）
- 引用内容是否准确
- 正文中的标点符号和格式
- 非中英文语言的引注

## 许可证

[MIT License](LICENSE)

## 参考

- [《法学引注手册》](https://www.press.tsinghua.edu.cn/) - 清华大学法学院编写
- 规则库基于手册第二版整理

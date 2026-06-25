# corrections.json 格式规范

`corrections.json` 是 `docx_revisor.py` 的输入文件，定义所有需要修订的脚注格式错误。

## JSON 结构

```json
[
  {
    "footnote_id": "3",
    "rule": "第96条",
    "error_type": "作者格式错误",
    "old_text": "Yan and Hyman",
    "new_text": "Yan & Hyman",
    "reason": "两位英文作者之间应使用 & 而非 and"
  },
  {
    "footnote_id": "3",
    "rule": "第103条",
    "error_type": "书名未用斜体",
    "old_text": "Introduction to Comparative Law",
    "new_text": "Introduction to Comparative Law",
    "reason": "英文书名应使用斜体",
    "format_changes": {
      "italic": true
    }
  }
]
```

## 字段说明

| 字段 | 类型 | 必填 | 说明 |
|------|------|------|------|
| `footnote_id` | string | 是 | 脚注编号（对应 `docx_parser.py` 输出的 `id` 字段） |
| `rule` | string | 是 | 违反的手册条款编号，如 `"第25条"` |
| `error_type` | string | 是 | 错误类型简述，如 `"缺少冒号"`、`"标点符号错误"` |
| `old_text` | string | 是 | 脚注中**实际存在**的原始文本片段 |
| `new_text` | string | 是 | 替换后的修正文本片段 |
| `reason` | string | 是 | 详细修改理由，会写入蓝色批注 |
| `comment_only` | bool | 否 | 仅添加批注而不修改文本（默认 `false`）。用于需要人工判断的场景 |
| `format_changes` | object | 否 | 格式修正设置，用于修正斜体/正体等格式属性。详见下方"格式修正"章节 |

## old_text 要求

`old_text` 必须满足以下条件：

1. **精确匹配**：在脚注文本中确实存在，逐字符一致。工具通过 XML 层的文本查找来定位。
2. **唯一性**：在所属脚注中唯一。如果同一文本出现多次，应选择包含上下文更长的唯一片段。
3. **最小粒度**：只包含需要替换的部分，不包含上下文无关的文本。
4. **实际内容**：基于 `docx_parser.py` 提取的 `text` 字段，而非文档中的原始排版。

## 典型修正示例

### 中文-缺"载"字

```json
{
  "footnote_id": "5",
  "rule": "第42条",
  "error_type": "缺少\"载\"字",
  "old_text": "》，《德国研究》",
  "new_text": "》，载《德国研究》",
  "reason": "文章引注缺少\"载\"字，应在来源期刊前加\"载\""
}
```

### 中文-页码格式

```json
{
  "footnote_id": "5",
  "rule": "第48条",
  "error_type": "页码格式错误",
  "old_text": "68-79页",
  "new_text": "第68-79页",
  "reason": "页码缺少\"第\"字，正确格式为\"第X页\""
}
```

### 中文-标点符号

```json
{
  "footnote_id": "15",
  "rule": "第24条",
  "error_type": "标点符号错误",
  "old_text": "孙谦.",
  "new_text": "孙谦：",
  "reason": "作者名后应使用全角冒号（：），而非英文句点（.）"
}
```

### 英文-作者连接词

```json
{
  "footnote_id": "3",
  "rule": "第96条",
  "error_type": "作者连接词错误",
  "old_text": " and ",
  "new_text": " & ",
  "reason": "两位英文作者之间应使用 & 而非 and"
}
```

### 英文-页码格式

```json
{
  "footnote_id": "1",
  "rule": "第107条",
  "error_type": "页码格式错误",
  "old_text": "pp.",
  "new_text": "p.",
  "reason": "英文引注无论单页还是多页，均使用 p. 而非 pp."
}
```

### 英文-书名斜体修正

```json
{
  "footnote_id": "3",
  "rule": "第103条",
  "error_type": "书名未用斜体",
  "old_text": "Introduction to Comparative Law",
  "new_text": "Introduction to Comparative Law",
  "reason": "英文书名应使用斜体，而非正体",
  "format_changes": {
    "italic": true
  }
}
```

### 英文-文章名去斜体（改回正体）

```json
{
  "footnote_id": "5",
  "rule": "第103条",
  "error_type": "文章名误用斜体",
  "old_text": "Hard Cases",
  "new_text": "Hard Cases",
  "reason": "英文文章标题应使用正体，不应斜体",
  "format_changes": {
    "italic": false
  }
}
```

### 纯批注（不修改文本）

```json
{
  "footnote_id": "12",
  "rule": "第96条",
  "error_type": "建议人工复核",
  "old_text": "",
  "new_text": "",
  "reason": "该引注含中英文混排，请人工确认标点符号使用是否正确",
  "comment_only": true
}
```

## 格式修正（`format_changes`）

用于修正字符格式属性，如斜体、正体等。当前支持的属性：

| 属性 | 类型 | 说明 |
|------|------|------|
| `italic` | bool | `true` = 应用斜体；`false` = 移除斜体改用正体 |

格式修正通过修订模式的 `w:ins` 元素实现：被标记为"删除"的旧文本保留原始斜体状态，被标记为"插入"的新文本应用 `format_changes` 指定的斜体状态。Word 接受修订后新文本将获得正确的格式。

**适用场景**（根据《法学引注手册》规则）：

| 规则 | 应使用斜体 | 应使用正体 |
|------|-----------|-----------|
| 第103条 | 书名（如 *Introduction to Comparative Law*） | 文章标题 |
| 第103条 | 文集名（如 *Taking Rights Seriously*） | — |
| 第112条 | 案例名（如 *Marbury v. Madison*） | — |
| 第112条 | v.（如 Marbury *v.* Madison） | — |
| 第106条 | — | 报纸名正体或斜体均可（全书统一） |

> **注意**：`old_text` 和 `new_text` 可以相同，此时仅修改格式不修改文字。修订模式会显示为"格式变更"。

## 生成注意事项

1. **按脚注编号排序**：先按 `footnote_id` 排序，同一脚注内按文本出现先后排序。
2. **避免重复替换**：同一脚注中不要有两条修正的 `old_text` 重叠或互为子串。
3. **测试 old_text 唯一性**：生成后应确认每条 `old_text` 在对应脚注的 `text` 字段中确实存在且唯一。
4. **错误优先于警告**：格式错误用 `comment_only: false`（修改+批注），不确定的建议用 `comment_only: true`（仅批注）。

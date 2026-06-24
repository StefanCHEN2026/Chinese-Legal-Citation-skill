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

## 生成注意事项

1. **按脚注编号排序**：先按 `footnote_id` 排序，同一脚注内按文本出现先后排序。
2. **避免重复替换**：同一脚注中不要有两条修正的 `old_text` 重叠或互为子串。
3. **测试 old_text 唯一性**：生成后应确认每条 `old_text` 在对应脚注的 `text` 字段中确实存在且唯一。
4. **错误优先于警告**：格式错误用 `comment_only: false`（修改+批注），不确定的建议用 `comment_only: true`（仅批注）。

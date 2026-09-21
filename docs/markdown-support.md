# Supported Markdown syntax

## Supported

- ATX heading: `#` ～ `####`
- Setext heading: level 1 / level 2
- Bullet list: `-`, `*`, `+`
- Numbered list: `1.` / `1)`
- Block quote and nested block quote
- Fenced code block: backtick / tilde
- GFM-style table
- Emphasis: `*`, `_`
- Strong emphasis: `**`, `__`
- Inline code with variable-length backticks
- Hard line break:
  - two trailing spaces
  - trailing backslash
  - `<br>`, `<br/>`, `<br />`
- Backslash escape for ASCII punctuation

## Not supported or intentionally different

- Links and images
- Autolinks
- HTML rendering other than `<br>`
- Strikethrough
- Task list
- Indented code block
- Ordered-list numbering normalization
- Table alignment in Excel
- Full CommonMark block parsing
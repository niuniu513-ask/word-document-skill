# Examples

This directory contains example `content.json` files demonstrating the content
model used by `generate_docx.py`.

## Files

| File | Description |
|---|---|
| `lab-report-content.json` | Full lab report with cover fields, headings, table_fill, new table, image, and page break |
| `notice-content.json` | Simple formal notice (title + body + signature) |

## Usage

```bash
# With a template (insert mode — least invasive)
python scripts/generate_docx.py --template template.docx \
    --content examples/lab-report-content.json -o output.docx

# From a document spec (new document)
python scripts/analyze_docx.py reference.docx -o spec.json
python scripts/generate_docx.py --spec spec.json \
    --content examples/notice-content.json -o output.docx
```

## Content Model Quick Reference

```json
{"blocks": [
  {"type": "cover", "fields": {"姓名": "张三"}},
  {"type": "heading", "level": 1, "text": "一、标题", "numbering": "native"},
  {"type": "paragraph", "text": "正文内容", "after": "一、标题"},
  {"type": "table", "rows": [["列1","列2"]], "caption": "表1"},
  {"type": "table_fill", "table_index": 0, "cells": {"1,0": "值"}},
  {"type": "image", "path": "fig.png", "width_cm": 10, "caption": "图1"},
  {"type": "caption", "text": "题注文本"},
  {"type": "page_break"}
]}
```

See `schemas/content-model.json` for the full JSON Schema definition.

# Markdown-to-Word Delivery Demo

## 1 Introduction

This sample is intentionally small and contains only public demonstration
content. It shows how a Markdown source, a reference DOCX, and a profile form a
repeatable delivery unit.

## 2 Captions and code

表2-1 Validation contract

| Gate | Expected result |
| --- | --- |
| Caption | Figure and table labels have stable numbering |
| Package | DOCX ZIP integrity passes |

```python
def build_report(source: str) -> str:
    return f"Built {source}"
```

图2-1 Delivery flow

The diagram can be supplied as a normal Markdown image in a real profile. This
public sample keeps the repository free of third-party visual assets.

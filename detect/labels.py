"""段落开头的「标签」识别：Abstract: / 摘要 / Keywords: / 关键词。

期刊论文习惯把这几个字加粗，正文部分不加粗。排版和预览都要用同一份规则，
所以单独放一个模块（`reflow` 写 docx、`render` 画预览都 import 它）。
"""
from __future__ import annotations

import re

# 注意顺序：长写法在前，避免 "Abstract:" 被 "Abstract" 截断后退成不匹配
RE_LABEL = re.compile(
    r"^\s*(Abstract|ABSTRACT|Abstract\s*[:：]|摘要\s*[:：]?|"
    r"Keywords?|KEYWORDS?|关键词)\s*[:：]?\s*",
    re.I,
)

# 只有这些角色的段落才做标签加粗（摘要 / 关键词）
LABEL_ROLES = ("abstract", "keywords")


def split_label(text: str):
    """返回 (标签, 其余)。不是标签开头就返回 (None, text)。

    标签必须短于整段，否则整段就是一个词（比如标题本身就叫 "Abstract"），
    那时没有"其余"部分，加粗整个词没有意义。
    """
    m = RE_LABEL.match(text or "")
    if not m:
        return None, text
    label = m.group(0)
    if not label or len(label) >= len(text):
        return None, text
    return label, text[len(label):]

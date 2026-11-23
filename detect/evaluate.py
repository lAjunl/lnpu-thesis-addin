"""识别准确率评测。

真值来自文档自身的样式（这份稿子有完整样式，等于白送标签）。
但"样式名 ≠ 语义角色"，所以同时报两套口径：
  A) 样式口径：预测角色 vs 样式直接映射的角色
  B) 语义口径：预测角色 vs 人工确认的语义映射（论文标题单独算 title）

用法:
    .venv/Scripts/python.exe detect/evaluate.py <docx路径> [--use-style]
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from collections import Counter, defaultdict
from pathlib import Path

from docx import Document

from .features import extract_features, body_baseline
from .classify import classify

# 样式名 → 角色（口径 A）
STYLE_TO_ROLE = {
    "heading 1": "heading1",
    "heading 2": "heading2",
    "heading 3": "heading3",
    "heading 4": "heading4",
    "image caption": "figure_caption",
    "captioned figure": "figure",
    "body text": "body",
    "first paragraph": "body",
    "compact": "body",
    "caption": "figure_caption",
}

# 样式名 → 语义角色（口径 B）：本稿把 Heading 1 用作论文标题，
# 因此章节整体下移一级：Heading 2 → 一级标题，Heading 3 → 二级标题。
SEMANTIC_FROM_STYLE = {
    "heading 1": "title",
    "heading 2": "heading1",
    "heading 3": "heading2",
    "heading 4": "heading3",
    "image caption": "figure_caption",
    "captioned figure": "figure",
    "body text": "body",
    "first paragraph": "body",
    "compact": "body",
    "caption": "figure_caption",
}


def role_of_style(name: str, mapping: dict[str, str]) -> str:
    n = name.strip().lower()
    if n in mapping:
        return mapping[n]
    if "heading" in n:
        m = re.search(r"(\d)", n)
        return f"heading{m.group(1)}" if m else "other"
    if "caption" in n:
        return "figure_caption"
    return "other"


def find_reference_zone(feats) -> int | None:
    """定位参考文献区起点。

    独立于样式：找"最长连续 [N] 递增序列"的起点。文献列表的编号必然连续，
    而正文里的零星引用不会形成长连续序列。
    """
    best_start = None
    best_len = 0
    i = 0
    n = len(feats)
    while i < n:
        m = re.match(r"^\s*\[(\d+)\]", feats[i].text)
        if not m:
            i += 1
            continue
        expect = int(m.group(1)) + 1
        j = i + 1
        length = 1
        while j < n:
            m2 = re.match(r"^\s*\[(\d+)\]", feats[j].text)
            if not m2 or int(m2.group(1)) != expect:
                break
            expect += 1
            j += 1
            length += 1
        if length > best_len:
            best_len = length
            best_start = i
        i = j if j > i else i + 1
    return best_start if best_len >= 10 else None


def semantic_truth(feats) -> tuple[list[str], dict]:
    """语义真值：以样式为基础，但用独立依据覆盖三类。

    覆盖依据（都不依赖分类器的输出，避免循环论证）：
      1. 文献区内连续 [N] 递增的段落 → reference
      2. 文本就是 "摘要/Abstract" → abstract
      3. 文本以 "关键词/Keywords" 开头 → keywords
    另外本文把 Heading 1 用作论文标题，章节整体下移一级。
    """
    ref_start = find_reference_zone(feats)
    out: list[str] = []
    for f in feats:
        role = role_of_style(f.style_name, SEMANTIC_FROM_STYLE)
        if ref_start is not None and f.index >= ref_start and RE_REF_LINE.match(f.text):
            role = "reference"
        elif RE_ABSTRACT_LINE.match(f.text):
            role = "abstract"
        elif RE_KEYWORD_LINE.match(f.text):
            role = "keywords"
        out.append(role)
    return out, {"ref_start": ref_start}


RE_REF_LINE = re.compile(r"^\s*\[\d+\]")
RE_ABSTRACT_LINE = re.compile(r"^\s*(摘\s*要|abstract)\s*$", re.I)
RE_KEYWORD_LINE = re.compile(r"^\s*(关键词|keywords)\s*[:：]", re.I)


def metrics(y_true: list[str], y_pred: list[str]) -> dict:
    n = len(y_true)
    correct = sum(1 for a, b in zip(y_true, y_pred) if a == b)
    per: dict[str, dict] = {}
    for role in sorted(set(y_true) | set(y_pred)):
        tp = sum(1 for a, b in zip(y_true, y_pred) if a == role and b == role)
        fp = sum(1 for a, b in zip(y_true, y_pred) if a != role and b == role)
        fn = sum(1 for a, b in zip(y_true, y_pred) if a == role and b != role)
        support = sum(1 for a in y_true if a == role)
        per[role] = {
            "support": support,
            "precision": tp / (tp + fp) if tp + fp else None,
            "recall": tp / (tp + fn) if tp + fn else None,
            "tp": tp,
            "fp": fp,
            "fn": fn,
        }
    return {"n": n, "accuracy": correct / n if n else 0.0, "per": per}


def print_report(title: str, y_true: list[str], y_pred: list[str], feats, show_fail: bool) -> dict:
    m = metrics(y_true, y_pred)
    print(f"\n===== {title} =====")
    print(f"总体准确率: {m['accuracy']:.1%}  ({m['n']} 段)")
    print(f"\n{'角色':<18}{'样本':>6}{'精确率':>9}{'召回率':>9}")
    for role, d in sorted(m["per"].items(), key=lambda x: -x[1]["support"]):
        p = f"{d['precision']:.0%}" if d["precision"] is not None else "—"
        r = f"{d['recall']:.0%}" if d["recall"] is not None else "—"
        print(f"{role:<18}{d['support']:>6}{p:>9}{r:>9}")

    # 混淆矩阵（只列出现过的组合）
    pairs = Counter(zip(y_true, y_pred))
    print("\n混淆（真值 → 预测 : 次数）:")
    for (t, p), c in pairs.most_common():
        flag = "  " if t == p else "❌"
        print(f"  {flag} {t:<18} → {p:<18} {c}")

    if show_fail:
        print("\n不一致样本（最多 30 条）:")
        n = 0
        for f, t, p in zip(feats, y_true, y_pred):
            if t == p:
                continue
            print(f"  #{f.index:<4d} 真={t:<16} 预={p:<16} style={f.style_name[:14]:<14} | {f.text[:44]}")
            n += 1
            if n >= 30:
                break
    return m


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("docx")
    ap.add_argument("--use-style", action="store_true", help="把样式名当线索（对照实验）")
    ap.add_argument("--fail", action="store_true", help="打印不一致样本")
    ap.add_argument("--truth", help="从 make_variants 产出的 ground_truth.json 载入真值")
    args = ap.parse_args()

    doc = Document(args.docx)
    feats = extract_features(doc)
    det = classify(feats, use_style=args.use_style)
    pred = [d.role for d in det]

    print(f"文件: {Path(args.docx).name}")
    print(f"段落: {len(feats)}   模式: {'使用样式线索' if args.use_style else '仅特征（模拟无格式稿）'}")
    base = body_baseline(feats)
    print(f"正文基线: 字号 {base['size']}pt  行距 {base['line_rule']} {base['line_value']}  对齐 {base['align']}  加粗 {base['bold']}  (样本 {base['n']})")

    if args.truth:
        data = json.loads(Path(args.truth).read_text(encoding="utf-8"))
        truth = [p["role"] for p in data["paragraphs"]]
        print(f"\n真值来源: {data['source']}（外部 ground_truth.json）")
        print("\n预测分布:", dict(Counter(pred)))
        print_report("外部真值口径", truth, pred, feats, args.fail)
        return

    style_true = [role_of_style(f.style_name, STYLE_TO_ROLE) for f in feats]
    sem_true, meta = semantic_truth(feats)
    print(f"\n语义真值修正: 文献区起点 = {meta['ref_start']}（连续 [N] 递增≥10 判定）")

    print("\n样式分布（真值）:", dict(Counter(f.style_name for f in feats)))
    print("预测分布:", dict(Counter(pred)))

    print_report("口径 A：样式恢复率（预测角色 vs 样式直接映射）", style_true, pred, feats, args.fail)
    print_report("口径 B：语义恢复率（文献区/摘要/关键词按内容判定，章节下移一级）", sem_true, pred, feats, args.fail)


if __name__ == "__main__":
    main()

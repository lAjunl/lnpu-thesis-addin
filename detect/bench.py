"""识别器基准测试。

用法:
    .venv/Scripts/python.exe -m detect.bench <docx> [--truth ground_truth.json] [--use-style] [--fail]

同时报告两个口径：
  - 标题识别率：预测为 heading（含未分级）即算认出了标题，只看"是不是标题"
  - 层级准确率：预测的 heading1/2/3 与真值层级完全一致
"""
from __future__ import annotations

import argparse
import json
import re
from collections import Counter
from pathlib import Path

from docx import Document

from .features import extract_features, body_baseline
from .classify import analyze
from .classify import classify


def is_heading(role: str) -> bool:
    return role == "heading" or (
        role.startswith("heading") and role != "heading"
    )


def level_of(role: str) -> int | None:
    m = re.fullmatch(r"heading(\d)", role or "")
    return int(m.group(1)) if m else None


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("docx")
    ap.add_argument("--truth")
    ap.add_argument("--use-style", action="store_true")
    ap.add_argument("--fail", action="store_true")
    args = ap.parse_args()

    doc = Document(args.docx)
    a = analyze(doc, source=str(args.docx), use_style=args.use_style)
    fs = a.features
    pred = [d.role for d in a.detections]

    print(f"文件: {Path(args.docx).name}   段落 {len(fs)}   "
          f"模式: {'含样式线索' if args.use_style else '仅特征'}")
    print(f"编号体系: {a.scheme}  {a.scheme_stats}")
    print(f"层级推断: {a.level_note}")
    b = body_baseline(fs)
    print(f"正文基线: 字号 {b['size']}pt  行距 {b['line_rule']}  样本 {b['n']}")
    print(f"预测分布: {dict(Counter(pred))}")
    if not args.truth:
        return

    data = json.loads(Path(args.truth).read_text(encoding="utf-8"))
    truth = [p["role"] for p in data["paragraphs"]]
    assert len(truth) == len(fs)

    strict = [(t, p) for t, p in zip(truth, pred) if t == p]
    loose = [
        (t, p)
        for t, p in zip(truth, pred)
        if t == p or (is_heading(t) and p == "heading")
    ]
    n = len(truth)
    print(f"\n严格准确率: {len(strict)/n:.1%}")
    print(f"标题识别率: {len(loose)/n:.1%}（未分级标题也算认出）")

    print(f"\n{'真值角色':<16}{'样本':>6}{'认出':>8}{'层级对':>8}")
    for role in sorted({t for t in truth}):
        idxs = [i for i, t in enumerate(truth) if t == role]
        found = sum(1 for i in idxs if pred[i] == role or (is_heading(role) and pred[i] == "heading"))
        lvl_ok = sum(1 for i in idxs if pred[i] == role)
        print(f"{role:<16}{len(idxs):>6}{found/len(idxs):>8.0%}{lvl_ok/len(idxs):>8.0%}")

    print("\n主要不一致（真 → 预 : 次数）:")
    for (t, p), c in Counter(
        (t, p) for t, p in zip(truth, pred) if t != p and not (is_heading(t) and p == "heading")
    ).most_common(10):
        print(f"  {t:<16} → {p:<16} {c}")

    if args.fail:
        print("\n不一致样本（最多 25 条）:")
        k = 0
        for f, t, p in zip(fs, truth, pred):
            if t == p or (is_heading(t) and p == "heading"):
                continue
            print(f"  #{f.index:<4d} 真={t:<14} 预={p:<14} {f.char_len:>4}字 {f.word_count:>3}词 | {f.text[:44]}")
            k += 1
            if k >= 25:
                break


if __name__ == "__main__":
    main()

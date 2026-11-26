"""生成排版预览文档：读原稿 → 重新识别 → 套格式/分节/编号 → 另存。

每次都从原稿重新读，保证预览干净（不累积上一次的效果）。
"""
from __future__ import annotations

from pathlib import Path

from docx import Document

from .classify import analyze, is_heading_role
from .numbering import (
    apply_plans,
    plan_caption_numbers,
    plan_heading_numbers,
)
from .reflow import apply_formats
from .sections import apply as apply_sections
from .sections import plans_from_settings


def derive_cut_points(analysis) -> tuple[int | None, int | None]:
    """前置起点 / 正文起点。"""
    front = None
    anchor = None
    body = None
    for d in analysis.detections:
        if front is None and d.role in ("abstract", "keywords", "toc"):
            front = d.index
            anchor = d.index
        if anchor is not None and body is None and is_heading_role(d.role) and d.index > anchor:
            body = d.index
    return front, body


def render_preview(source: Path, settings, out_path: Path,
                   overrides: dict | None = None,
                   analysis=None) -> tuple[int, int, int]:
    """返回 (排版段数, 编号数, 节数)。"""
    d = Document(str(source))
    a = analysis if analysis is not None else analyze(d, source=source.name)

    front = body = None
    if getattr(settings, "profile", "thesis") == "thesis":
        front, body = derive_cut_points(a)
        if front is not None and body is not None and (
            settings.enable_sections or settings.enable_page_numbers
        ):
            apply_sections(d, front_index=front, body_index=body,
                           plans=plans_from_settings(settings), settings=settings)
        else:
            from .sections import configure

            configure(d, plans_from_settings(settings), settings,
                      front_index=front, body_index=body)
    else:
        from .sections import configure

        configure(d, plans_from_settings(settings), settings)

    from .reflow import indent_chars_of

    stats = apply_formats(d, a, settings.roles, overrides=overrides,
                          label_bold=bool(getattr(settings, "bold_label", False)),
                          indent_chars=indent_chars_of(settings))
    from .reflow import style_tables
    style_tables(d, settings)
    if getattr(settings, "enable_marks", True):
        from .marks import apply_marks
        apply_marks(d)

    plans = []
    if settings.enable_number_heading:
        plans += plan_heading_numbers(a.detections, body, settings)
    if settings.enable_number_figure:
        plans += plan_caption_numbers(a.detections, "figure", settings, a.features)
    if settings.enable_number_table:
        plans += plan_caption_numbers(a.detections, "table", settings, a.features)
    n_num = apply_plans(d, plans) if plans else 0

    d.save(str(out_path))
    return sum(stats.values()), n_num, len(d.sections)

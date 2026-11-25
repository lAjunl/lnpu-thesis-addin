"""结构识别器：把段落判定为语义角色，并给出置信度与层级。

两种模式：
- use_style=True  样式名作为强线索（有样式的文档）
- use_style=False 只看格式/文本/位置特征（"无格式手写稿"的模拟）

编号体系（自动判断，判断不了交人工）：
- tech：1 / 1.1 / 1.1.1          → 层级 = 点数
- cn  ：一、 / （一） / 1. / （1） → 层级 1/2/3/4（第 4 级压到 3 级）

角色：title / heading1..3 / heading(未分级) / body / figure /
      figure_caption / table / table_caption / equation / reference /
      abstract / keywords / toc / meta
"""
from __future__ import annotations

import json
import re
import statistics
from dataclasses import dataclass, field
from pathlib import Path

from .features import ParaFeatures, body_baseline, extract_features

# ------------------------------------------------------------------ #
# 文本规则
# ------------------------------------------------------------------ #

RE_ABSTRACT = re.compile(r"^\s*(摘\s*要|abstract)\s*$", re.I)
# 期刊的 "Abstract: 正文…" ——整段就是摘要，标签后面还有内容
RE_ABSTRACT_LABEL = re.compile(r"^\s*(摘\s*要|abstract)\s*[:：]", re.I)
# Highlights / Graphical abstract：前置部分的小标题，不是章节，不能占层级，
# 否则「引言」会被编成 2（实测 test.docx 的 Highlights 被猜成一级标题）。
RE_HIGHLIGHTS = re.compile(
    r"^\s*(highlights?|research\s+highlights?|graphical\s+abstract|"
    r"highlights?\s+of\s+(the\s+)?(article|paper|manuscript|work))"
    r"\s*[:：]?\s*$", re.I)
RE_KEYWORDS = re.compile(r"^\s*(关键词|keywords)\s*[:：]", re.I)
RE_TOC_LINE = re.compile(r"^\s*(目\s*录|contents)\s*$", re.I)
RE_TOC_STYLE = re.compile(r"^\s*(?:toc|目录)\s*(\d+)\s*$", re.I)
RE_TOC_ENTRY = re.compile(r"(?:\t|[.·…]{3,})\s*\d+\s*$")

# 尾注区：不是章节，不能占章号（否则「参考文献」会变成第 N 章）
RE_BACK = re.compile(
    r"^\s*(references?|bibliography|funding|fund\s+support|grant\s+support|"
    r"declarations?|data\s+availability|acknowledg?e?ments?|"
    r"conflict\s+of\s+interest|author\s+contributions|appendix|appendices|"
    r"参考文献|致谢|基金|基金项目|声明|附录|作者贡献|利益冲突|数据可用性)"
    r"\s*[:：]?\s*$",
    re.I,
)

# 图注/表注：编号后面跟的是"标题式"的大写/中文，而不是小写动词（句子）
RE_CAP_NUM = re.compile(
    r"^\s*(?P<kind>图|表|Figure|Fig\.?|Table)\s*\d+\s*"
    r"[.．:：)】\](\-]*\s*(?P<next>\S)",
    re.I,
)
RE_CAP_PANEL = re.compile(
    r"^\s*(图|Figure|Fig\.?)\s*[.．:：]?\s*[（(]\s*[A-Za-z]\s*[)）]", re.I)


def _caption_kind(text: str) -> str | None:
    """判断是不是图注/表注开头。返回 "figure_caption" / "table_caption" / None。

    三种写法都要认（manuscript 实测 11 个图注只认出 1 个）：
        Fig. 1 Effect of …
        Fig.  Keyword co-occurrence …        （没编号，直接标题）
        Fig.  (a) CCUS process flow …        （分图）
    判别依据：编号词后面跟的是大写/中文（标题式），不是小写动词（句子）；
    且整段不是以句末标点结尾。
    """
    s = text.lstrip()
    m = re.match(r"^(图|表|Figure|Fig\.?|Table)\b", s, re.I)
    if not m:
        return None
    is_tab = m.group(1).lower().startswith(("表", "table"))
    rest = s[m.end():]

    m2 = re.match(r"^\s*\d+(?:\s*[.\-–—]\s*\d+)*\s*", rest)
    if m2:
        rest = rest[m2.end():]
        has_num = True
    else:
        has_num = False
    rest = re.sub(r"^[.．:：)】\](\-]*\s*", "", rest)
    if not rest:
        return None

    ch = rest[0]
    title_like = ch.isupper() or ("一" <= ch <= "鿿")
    if is_fig_kind(m.group(1)) and not has_num \
            and re.match(r"^[（(]\s*[A-Za-z]\s*[)）]", rest):
        return "figure_caption"
    # 带编号时靠"编号后面是大写/中文"区分句子和标题
    # （"Fig. 1 shows …" 小写 → 句子；"Fig. 1 Effect …" 大写 → 图注）
    if title_like:
        return "table_caption" if is_tab else "figure_caption"
    return None


def is_fig_kind(word: str) -> bool:
    return not word.lower().startswith(("表", "table"))

# 图注/表注强特征：编号后带点号，区别于正文里的引用（"Figure 1 shows…"）
RE_FIG_STRONG = re.compile(r"^\s*(图|figure|fig\.?)\s*\d+\s*[.、．]", re.I)
RE_TAB_STRONG = re.compile(r"^\s*(表|table)\s*\d+\s*[.、．]", re.I)
RE_FIG_WEAK = re.compile(r"^\s*(图|figure|fig\.?)\s*\d+", re.I)
RE_TAB_WEAK = re.compile(r"^\s*(表|table)\s*\d+", re.I)

RE_EQN = re.compile(r"^[（(]\s*式\s*\d+(?:\.\d+)?\s*[)）]")
# 公式编号，如 "(1)"、"（4-1）"、"（A.2）"
RE_EQ_LABEL = re.compile(r"^[（(]\s*[\w.\-–—]{1,12}\s*[)）][.。]?$")
RE_REF = re.compile(r"^\s*\[\d+\]")
RE_REF_ZONE = re.compile(r"^\s*(references|参考文献)\s*$", re.I)

# 封面元数据行（会被"短 + 非句末标点"误判成标题）
RE_META = re.compile(
    r"^\s*[-*•·†‡]?\s*(running\s+title|authors?|affiliations?|corresponding\s+author|"
    r"e-?mail|ORCID|journal|received|accepted|published|doi\s*[:：]|"
    r"摘\s*要|关键词|作\s*者|通\s*信|单\s*位|学\s*号)\s*[:：]?",
    re.I,
)

# ------------------------------------------------------------------ #
# 编号
# ------------------------------------------------------------------ #

# 科技体系：1 / 1.1 / 1.1.1（排除 2026年11月 / 15% 等年份/数量行）
RE_TECH_NUM = re.compile(
    r"^\s*(\d{1,2}(?:\.\d{1,2})+)\s*[.、．]?\s*(?:\S|$)"
    r"|^\s*(\d{1,2})(?:[.、．]\s*|\s+)(?![年月日%℃])(?:\S|$)"
)

# 公文式：一、 / （一） / 1. / （1）
RE_CN_L1 = re.compile(r"^\s*([一二三四五六七八九十百零]+)\s*、")
RE_CN_L2 = re.compile(r"^\s*[（(]\s*([一二三四五六七八九十百零]+)\s*[）)]")
RE_CN_L3 = re.compile(r"^\s*(\d{1,2})\s*[.．]\s*(?:\S|$)")
RE_CN_L4 = re.compile(r"^\s*[（(]\s*(\d{1,2})\s*[）)]")

CN_NUM = {"零": 0, "一": 1, "二": 2, "三": 3, "四": 4, "五": 5, "6": 6,
          "七": 7, "八": 8, "九": 9, "十": 10}


def detect_scheme(feats: list[ParaFeatures]) -> tuple[str, dict]:
    """判断全文用哪套编号体系。返回 (scheme, 统计)。"""
    cn1 = sum(1 for f in feats if RE_CN_L1.match(f.text))
    cn2 = sum(1 for f in feats if RE_CN_L2.match(f.text))
    cn4 = sum(1 for f in feats if RE_CN_L4.match(f.text))
    tech = sum(1 for f in feats if RE_TECH_NUM.match(f.text))
    stats = {"cn_l1": cn1, "cn_l2": cn2, "cn_l4": cn4, "tech": tech}
    # 出现汉字编号且明显多于括号数字 → 公文体系
    if cn1 + cn2 >= 3 and (cn1 + cn2) >= cn4:
        return "cn", stats
    return "tech", stats


def number_level(f: ParaFeatures, scheme: str) -> tuple[int | None, str | None]:
    """返回 (层级, 编号文本)。识别不出返回 (None, None)。"""
    if scheme == "cn":
        m = RE_CN_L1.match(f.text)
        if m:
            return 1, m.group(0).strip()
        m = RE_CN_L2.match(f.text)
        if m:
            return 2, m.group(0).strip()
        m = RE_CN_L3.match(f.text)
        if m:
            return 3, m.group(0).strip()
        m = RE_CN_L4.match(f.text)
        if m:
            return 3, m.group(0).strip()  # 第 4 级压到 3 级（学校最多三级）
        return None, None
    m = RE_TECH_NUM.match(f.text)
    if m:
        prefix = m.group(1) or m.group(2)
        lvl = min(prefix.count(".") + 1, 3)
        return lvl, prefix
    return None, None


# ------------------------------------------------------------------ #
# 层级启发
# ------------------------------------------------------------------ #

def is_heading_role(role: str) -> bool:
    return role == "heading" or re.fullmatch(r"heading[123]", role) is not None


def effective_role(d: "Detection") -> str:
    """把"未定级标题 + 层级"合成实际角色。

    识别器对无编号标题给出 role="heading"、level=1/2/3；
    界面和重排都要按"一级标题/二级标题"分组，所以这里统一。
    """
    if d.role == "heading" and d.level in (1, 2, 3):
        return f"heading{d.level}"
    return d.role


# 层级启发的阈值。来自 556 段英文综述的实测分布：
#   一级标题字数中位 15（10~42），二级中位 32（12~66）
#   一级词数中位 2（1~6），二级中位 5（2~8）
#   一级与前一个标题的间距中位 10，二级中位 5（但范围重叠严重）
LEVEL_SHORT_CHARS = 20
LEVEL_SHORT_WORDS = 3
LEVEL_BIG_GAP = 10


def guess_levels(
    feats: list[ParaFeatures],
    dets: list["Detection"],
    scheme: str,
) -> tuple[dict[int, int], dict[int, float], str]:
    """给未分级标题猜层级。

    有编号的标题是锚点，直接定级；其余用"字数 + 词数 + 间距"组合评分。
    间距单独用区分力不足（一级 gap 中位 10、二级 5，但范围重叠），
    所以只作为三个信号之一。

    返回 (levels, confidences, note)
    """
    headings = [i for i, d in enumerate(dets) if is_heading_role(d.role)]
    if not headings:
        return {}, {}, "无标题"

    gap: dict[int, int] = {}
    prev = -1
    for i in headings:
        gap[i] = (i - prev) if prev >= 0 else i + 1
        prev = i

    anchored: dict[int, int] = {}
    for i in headings:
        lv, _ = number_level(feats[i], scheme)
        if lv is not None:
            anchored[i] = lv

    levels: dict[int, int] = {}
    conf: dict[int, float] = {}
    for i in headings:
        if i in anchored:
            levels[i] = anchored[i]
            conf[i] = 0.9
            continue
        f = feats[i]
        score = 0
        if f.char_len <= LEVEL_SHORT_CHARS:
            score += 1
        if f.word_count <= LEVEL_SHORT_WORDS:
            score += 1
        if gap[i] >= LEVEL_BIG_GAP:
            score += 1
        levels[i] = 1 if score >= 2 else 2
        conf[i] = 0.5 if anchored else 0.35

    n_anchor = len(anchored)
    note = (
        f"编号锚点 {n_anchor} 个，其余按字数/词数/间距评分推断（建议，需人工确认）"
        if n_anchor
        else "无编号锚点，全部为评分推断（把握低，必须人工确认）"
    )
    return levels, conf, note


# ------------------------------------------------------------------ #
# 数据结构
# ------------------------------------------------------------------ #

@dataclass
class Detection:
    index: int
    role: str
    confidence: float
    reasons: list[str] = field(default_factory=list)
    level: int | None = None
    level_confidence: float = 0.0
    number: str | None = None
    user_confirmed: bool = False   # 用户已在界面上确认过层级/角色


@dataclass
class Analysis:
    source: str
    scheme: str
    scheme_stats: dict
    paragraph_count: int
    detections: list[Detection]
    features: list[ParaFeatures]
    level_note: str = ""

    def to_json(self, path: str | Path) -> None:
        data = {
            "source": self.source,
            "numbering_scheme": self.scheme,
            "scheme_stats": self.scheme_stats,
            "paragraph_count": self.paragraph_count,
            "level_note": self.level_note,
            "summary": self.summary(),
            "paragraphs": [
                {
                    "index": d.index,
                    "text": self.features[d.index].text,
                    "role": d.role,
                    "level": d.level,
                    "level_confidence": round(d.level_confidence, 2),
                    "confidence": round(d.confidence, 2),
                    "number": d.number,
                    "reasons": d.reasons,
                }
                for d in self.detections
            ],
        }
        Path(path).write_text(json.dumps(data, ensure_ascii=False, indent=1), encoding="utf-8")

    def summary(self) -> dict:
        from collections import Counter

        c = Counter(d.role for d in self.detections)
        heads = sum(v for k, v in c.items() if is_heading_role(k))
        return {
            "paragraphs": self.paragraph_count,
            "headings": heads,
            "by_role": dict(c),
        }


# ------------------------------------------------------------------ #
# 主流程
# ------------------------------------------------------------------ #

def _norm_style(name: str) -> str:
    n = name.strip().lower().replace("_", " ")
    if RE_TOC_STYLE.match(n):
        return "toc_item"
    if "heading" in n or "标题" in n:
        m = re.search(r"(\d+)", n)
        if m:
            lvl = min(max(int(m.group(1)), 1), 3)
            return f"heading{lvl}"
        return "heading"
    if "caption" in n:
        if "figure" in n or "图" in n:
            return "figure_caption"
        if "table" in n or "表" in n:
            return "table_caption"
        return "caption"
    if "title" in n:
        return "title"
    if any(k in n for k in ("body", "first paragraph", "normal", "正文", "compact")):
        return "body"
    return "other"


def classify(
    feats: list[ParaFeatures],
    use_style: bool = False,
    scheme: str = "tech",
) -> list[Detection]:
    base = body_baseline(feats)
    body_size = base["size"]
    body_bold = base["bold"]
    n = len(feats)
    # 第一个非空段：题目候选（第 0 段可能是空段）
    first_nonempty = next((f.index for f in feats if f.char_len > 0), None)

    ref_start = None
    for f in feats:
        if RE_REF_ZONE.match(f.text):
            ref_start = f.index
            break

    out: list[Detection] = []
    ref_seen = 0

    for f in feats:
        # ---------- 硬规则（与格式无关） ----------
        if f.has_image:
            out.append(Detection(f.index, "figure", 0.99, ["段落内含图片"]))
            continue
        if f.in_table:
            out.append(Detection(f.index, "table", 0.95, ["位于表格内"]))
            continue
        if f.has_math or f.has_ole:
            # 只有"公式独占一段"才算公式段。公式夹在句子里（行内公式）时，
            # 这段仍然是正文——否则会被居中、换行距，Word 里整段都排错。
            rest = (f.text or "").replace("\t", " ").strip()
            if not rest or RE_EQ_LABEL.match(rest):
                what = "OMML" if f.has_math else "MathType"
                out.append(Detection(f.index, "equation", 0.95,
                                     [f"段落只含公式对象（{what}）"]))
                continue
        # 空段落永远不是标题：旧规则 "字数≤65 且词数≤12" 把 0 字也放进来，
        # 实测 manuscript 里 8 个空段落被编成一级标题、各占一个章号。
        if f.char_len == 0:
            out.append(Detection(f.index, "other", 0.9, ["空段落"]))
            continue
        if RE_EQN.match(f.text):
            out.append(Detection(f.index, "equation", 0.9, ["公式序号"]))
            continue

        if RE_REF.match(f.text):
            in_ref_zone = ref_start is not None and f.index > ref_start
            adjacent = ref_seen >= 3 and f.char_len >= 40
            if in_ref_zone or adjacent:
                ref_seen += 1
                out.append(Detection(f.index, "reference", 0.85, ["文献区内的 [N] 条目"]))
                continue
            if f.char_len >= 40 and f.index > n * 0.6:
                ref_seen += 1
                out.append(Detection(f.index, "reference", 0.55, ["后 40% 的 [N] 长段落"]))
                continue

        if RE_TOC_LINE.match(f.text):
            out.append(Detection(f.index, "toc", 0.9, ["目录标题"]))
            continue
        m_toc = RE_TOC_STYLE.match(f.style_name) if f.style_name else None
        if m_toc or f.is_toc_field:
            if m_toc:
                tlvl = min(max(int(m_toc.group(1)), 1), 3)
            else:
                lv_t, _ = number_level(f, scheme)
                tlvl = lv_t or 1
            why = [f"目录条目（样式 {f.style_name}）" if m_toc else "目录条目（Word TOC 域）"]
            out.append(Detection(f.index, "toc_item", 0.95, why, level=tlvl,
                                 level_confidence=0.95))
            continue
        if RE_BACK.match(f.text):
            out.append(Detection(f.index, "back", 0.85,
                                 ["尾注区标题（不参与章号）"]))
            continue
        if RE_KEYWORDS.match(f.text):
            out.append(Detection(f.index, "keywords", 0.9, ["关键词行"]))
            continue
        if RE_ABSTRACT.match(f.text):
            out.append(Detection(f.index, "abstract", 0.85, ["摘要标题"]))
            continue
        if RE_ABSTRACT_LABEL.match(f.text):
            out.append(Detection(f.index, "abstract", 0.85,
                                 ["摘要段（Abstract: 开头）"]))
            continue
        if RE_HIGHLIGHTS.match(f.text):
            out.append(Detection(f.index, "abstract", 0.8,
                                 ["前置部分标题（Highlights）"]))
            continue

        if RE_FIG_STRONG.match(f.text):
            out.append(Detection(f.index, "figure_caption", 0.88, ["“图/Figure N.”带点号"]))
            continue
        if RE_TAB_STRONG.match(f.text):
            out.append(Detection(f.index, "table_caption", 0.88, ["“表/Table N.”带点号"]))
            continue
        # 弱特征不再要求"紧邻图片"：表注在表格上方、附近没有图片，
        # 旧规则会漏掉 "Table 2 Elemental compositions…"，被猜成二级标题。
        # 用"编号后面跟的是大写/中文（标题式），不是小写动词（句子）"来区分。
        cap = _caption_kind(f.text)
        if cap and f.char_len <= 600:
            out.append(Detection(f.index, cap, 0.72,
                                 ["“图/Table N”短行（非句末标点）"]))
            continue
        near_img = (f.prev_image_idx is not None and f.index - f.prev_image_idx <= 3) or (
            f.next_image_idx is not None and f.next_image_idx - f.index <= 3
        )
        if near_img and f.char_len <= 90 and not f.ends_with_punct:
            if RE_FIG_WEAK.match(f.text):
                out.append(Detection(f.index, "figure_caption", 0.6, ["紧邻图片的“图 N”短行"]))
                continue
            if RE_TAB_WEAK.match(f.text):
                out.append(Detection(f.index, "table_caption", 0.6, ["紧邻图片的“表 N”短行"]))
                continue

        short = f.char_len <= 70
        no_end_punct = not f.ends_with_punct
        size = f.size_pt
        bigger = size is not None and body_size is not None and size > body_size + 0.5
        bolder = f.bold is True and body_bold is False

        # 首段（以及紧跟在首段后面的多行大号居中封面标题）优先于样式判成题目：
        # 很多工程报告/论文的封面标题用 Normal 样式手动放大字号，不能被 Normal 吞成正文
        if f.index == first_nonempty and (
                bigger or bolder or f.align == "center" or f.word_count <= 15):
            out.append(Detection(f.index, "title", 0.7, ["首段且字号/加粗/居中突出"]))
            continue
        if f.index == first_nonempty and f.char_len >= 40 and not f.ends_with_punct:
            out.append(Detection(f.index, "title", 0.6,
                                 [f"首段且 {f.char_len} 字、无句末标点"]))
            continue
        if (out and out[-1].role == "title" and f.index <= (first_nonempty or 0) + 3
                and size is not None and feats[out[-1].index].size_pt is not None
                and abs(size - feats[out[-1].index].size_pt) <= 0.5
                and size >= 18.0 and f.align == "center" and no_end_punct):
            out.append(Detection(f.index, "title", 0.7, ["多行封面主标题续行"]))
            continue

        # ---------- 有样式线索 ----------
        if use_style and f.style_name:
            ns = _norm_style(f.style_name)
            if ns in ("heading1", "heading2", "heading3"):
                lvl = int(ns[-1])
                _, num = number_level(f, scheme)
                out.append(Detection(f.index, ns, 0.95, [f"样式 {f.style_name}"],
                                     level=lvl, number=num, level_confidence=0.95))
                continue
            if ns in ("figure_caption", "table_caption", "caption"):
                out.append(Detection(f.index, ns, 0.9, [f"样式 {f.style_name}"]))
                continue
            if ns == "title":
                if f.index <= 15:
                    out.append(Detection(f.index, "title", 0.85, [f"样式 {f.style_name}"]))
                else:
                    out.append(Detection(f.index, "other", 0.75, [f"非封面区的 {f.style_name} 样式"]))
                continue
            if RE_META.match(f.text):
                out.append(Detection(f.index, "meta", 0.7, ["封面元数据行"]))
                continue
            # 文档本身已经规范使用 Heading 样式时，其余 Normal/自定义正文样式不再靠"短行"猜标题，
            # 避免把缩略语表、标准清单误判成上百个标题
            out.append(Detection(f.index, "body", 0.85, [f"样式 {f.style_name}"]))
            continue

        # ---------- 特征推断（仅无样式稿走此分支） ----------
        if RE_META.match(f.text):
            out.append(Detection(f.index, "meta", 0.7, ["封面元数据行"]))
            continue

        if short and no_end_punct:
            lv, num = number_level(f, scheme)
            if lv is not None:
                why = [f"编号 {num}（{scheme} 体系）"]
                if bigger:
                    why.append(f"字号 {size}pt > 正文 {body_size}pt")
                if bolder:
                    why.append("加粗")
                out.append(Detection(f.index, f"heading{lv}", 0.82, why, level=lv, number=num,
                                     level_confidence=0.9))
                continue
            if bigger or bolder:
                lvl = 1 if (size is not None and body_size is not None and size > body_size + 2) else 2
                out.append(Detection(f.index, f"heading{lvl}", 0.6, [f"字号 {size}pt 突出"],
                                     level=lvl, level_confidence=0.5))
                continue
            if f.char_len <= 65 and f.word_count <= 12:
                out.append(
                    Detection(
                        f.index,
                        "heading",
                        0.55,
                        [f"短（{f.char_len} 字 / {f.word_count} 词）、非句末标点，层级未定"],
                    )
                )
                continue

        out.append(Detection(f.index, "body", 0.6 if short else 0.85, ["无突出特征"]))

    _mark_toc_entries(out, feats, scheme)
    _mark_cover_meta(out, feats)
    _demote_unnumbered_headings(out)
    _mark_abstract_body(out, feats)
    return out


def _mark_toc_entries(out: list["Detection"], feats: list[ParaFeatures],
                      scheme: str) -> None:
    """无样式稿里的手工目录：紧跟在「目 录」后面、以制表符/点线+页码结尾的行标为 toc_item。"""
    n = len(out)
    i = 0
    while i < n:
        if out[i].role != "toc":
            i += 1
            continue
        j = i + 1
        while j < n:
            f = feats[out[j].index]
            if f.char_len == 0:
                j += 1
                continue
            if out[j].role == "toc_item":
                j += 1
                continue
            if RE_TOC_ENTRY.search(f.text) and f.char_len <= 100:
                lv, _ = number_level(f, scheme)
                out[j] = Detection(f.index, "toc_item", 0.9,
                                   ["目录标题之后的页码行"], level=lv or 1,
                                   level_confidence=0.85)
                j += 1
                continue
            break
        i = j


def _demote_unnumbered_headings(out: list["Detection"]) -> int:
    """原文标题普遍带编号时，没编号的短行按正文处理。

    无样式、无编号的稿子只能靠"短 + 非句末标点"猜标题，会把
    「Overall anodic dissolution reaction」这种插入语猜成二级标题，
    把后面的 3.2/3.3 整体顶掉一位。原文里真标题几乎都带编号时，
    没编号的短行就不该占层级。
    """
    heads = [d for d in out if is_heading_role(d.role)]
    if len(heads) < 4:
        return 0
    numbered = sum(1 for d in heads if d.number)
    if numbered / len(heads) < 0.6:
        return 0
    n = 0
    for i, d in enumerate(out):
        if not is_heading_role(d.role) or d.number:
            continue
        if d.confidence > 0.75:      # 样式/字号等硬证据不动
            continue
        out[i] = Detection(d.index, "body", 0.55,
                           list(d.reasons) + ["原文标题普遍带编号，此行无编号"])
        n += 1
    return n


_META_HINT = re.compile(
    r"[@＠]|universit|college|institute|institut|department|school|academy|"
    r"author|affiliation|corresponding|e-?mail|orcid|"
    r"大学|学院|系|专业|导师|作者|单位|通讯|邮箱",
    re.I,
)


def _looks_like_meta(text: str) -> bool:
    """作者/单位行：带典型单位词，或"逗号分隔的多个姓名"。"""
    if _META_HINT.search(text):
        return True
    return text.count(",") >= 2 and not re.match(r"^\s*\d", text)


def _mark_cover_meta(out: list["Detection"], feats: list[ParaFeatures],
                     limit: int = 8) -> None:
    """题目之后、摘要/关键词/第一个章节标题之前的信息行 → meta。

    期刊论文的作者/单位/通讯常常没有 "Authors:" 前缀，只是一行短行，
    旧规则会把它猜成标题（实测 manuscript 第 1 段被猜成二级标题）。
    """
    title_i = next((i for i, d in enumerate(out) if d.role == "title"), None)
    if title_i is None:
        return
    seen = 0
    for j in range(title_i + 1, min(len(out), title_i + 35)):
        d = out[j]
        if d.role == "title":
            continue
        if d.role in ("abstract", "keywords", "toc", "toc_item", "back",
                      "figure", "table", "equation", "reference", "meta"):
            break                      # 遇到真正的结构元素就停
        f = feats[d.index]
        if f.char_len == 0:
            continue
        if d.role == "other":
            break
        seen += 1
        if seen > limit:
            break
        if f.char_len > 150 or f.ends_with_punct:
            break
        if d.role == "body":
            out[j] = Detection(d.index, "meta", 0.65,
                               ["题目之后的作者/单位信息行"])
            continue
        if is_heading_role(d.role):
            # 已经被猜成标题的作者行：带编号的是真标题，否则按元数据处理
            if re.match(r"^\s*\d", f.text) or not _looks_like_meta(f.text):
                break
            out[j] = Detection(d.index, "meta", 0.6,
                               ["题目之后的作者/单位信息行"])


def _mark_abstract_body(out: list[Detection], feats: list[ParaFeatures] | None = None,
                        max_paras: int = 12) -> None:
    """把「摘要」标题之后、到下一个关键词/标题为止的正文段标成 abstract_body。

    学校规范普遍对摘要正文单独规定字号和行距，但以前的识别只认「摘要」那一行，
    正文段落一律落到 body，于是没法单独设。

    期刊的摘要是「Abstract: 正文…」一整段，后面紧跟着就是正文，
    这种不能再往下吞（否则整篇正文都变成摘要正文）。
    """
    n = len(out)
    i = 0
    while i < n:
        if out[i].role != "abstract":
            i += 1
            continue
        if feats is not None and feats[out[i].index].char_len > 80:
            i += 1
            continue
        j = i + 1
        taken = 0
        while j < n and taken < max_paras and out[j].role == "body":
            d = out[j]
            out[j] = Detection(d.index, "abstract_body", 0.8,
                               ["摘要标题之后的正文"])
            taken += 1
            j += 1
        i = j if j > i + 1 else i + 1


def _has_heading_styles(feats: list[ParaFeatures]) -> bool:
    """文档本身是否规范使用了 Word 的 Heading / 标题 样式（至少 3 个非空段）。"""
    n = sum(1 for f in feats if f.char_len > 0 and f.style_name
            and _norm_style(f.style_name) in ("heading1", "heading2", "heading3"))
    return n >= 3


def analyze(
    doc,
    source: str = "",
    use_style: bool | None = None,
) -> Analysis:
    feats = extract_features(doc)
    scheme, stats = detect_scheme(feats)
    if use_style is None:
        use_style = _has_heading_styles(feats)
    dets = classify(feats, use_style=use_style, scheme=scheme)

    levels, confs, note = guess_levels(feats, dets, scheme)
    for d in dets:
        if d.level is None and d.index in levels:
            d.level = levels[d.index]
            d.level_confidence = confs[d.index]
        elif d.level is not None and d.index not in levels:
            d.level_confidence = 0.9

    return Analysis(
        source=source,
        scheme=scheme,
        scheme_stats=stats,
        paragraph_count=len(feats),
        detections=dets,
        features=feats,
        level_note=note,
    )

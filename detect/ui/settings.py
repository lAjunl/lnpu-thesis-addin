"""格式设置数据模型：单一数据源。

默认值 = 学校学位论文格式的通用取值（可按需改）。
界面、预览、排版引擎都读这一份。
"""
from __future__ import annotations

from dataclasses import dataclass, field, asdict

# ------------------------------------------------------------------ #
# 学校规范默认值
# ------------------------------------------------------------------ #

# 页边距（mm）
PAGE_DEFAULT = {
    "top": 25.0, "bottom": 25.0, "left": 30.0, "right": 30.0,
    "header": 15.0, "footer": 20.0,
}
PAPER = "A4"
LINE_SPACING_PT = 22.0
HEADER_TEXT = "硕士学位论文"

# 角色默认格式
ROLE_DEFAULTS = {
    "title": {"size": 16.0, "font": "宋体", "ascii": "Times New Roman",
              "bold": True, "line": 22.0, "align": "center",
              "line_rule": "exact", "color": ""},
    "heading1": {"size": 15.0, "font": "宋体", "ascii": "Times New Roman",
                 "bold": True, "line": 22.0, "align": "left",
                 "line_rule": "exact", "color": ""},
    "heading2": {"size": 14.0, "font": "宋体", "ascii": "Times New Roman",
                 "bold": False, "line": 22.0, "align": "left",
                 "line_rule": "exact", "color": ""},
    "heading3": {"size": 12.0, "font": "宋体", "ascii": "Times New Roman",
                 "bold": False, "line": 22.0, "align": "left",
                 "line_rule": "exact", "color": ""},
    "body": {"size": 12.0, "font": "宋体", "ascii": "Times New Roman",
             "bold": False, "line": 22.0, "align": "both",
             "line_rule": "exact", "color": ""},
    "abstract_body": {"size": 12.0, "font": "宋体", "ascii": "Times New Roman",
                      "bold": False, "line": 22.0, "align": "both",
                      "line_rule": "exact", "color": ""},
    "figure_caption": {"size": 10.5, "font": "宋体", "ascii": "Times New Roman",
                       "bold": True, "line": 22.0, "align": "center",
                       "line_rule": "exact", "color": ""},
    "table_caption": {"size": 10.5, "font": "宋体", "ascii": "Times New Roman",
                      "bold": True, "line": 22.0, "align": "center",
                      "line_rule": "exact", "color": ""},
    # 未定级标题：按一级处理
    "heading": {"size": 15.0, "font": "宋体", "ascii": "Times New Roman",
                "bold": True, "line": 22.0, "align": "left",
                "line_rule": "exact", "color": ""},
    # 前置部分（这些角色以前没有格式，排版时完全不生效）
    "abstract": {"size": 14.0, "font": "宋体", "ascii": "Times New Roman",
                 "bold": True, "line": 22.0, "align": "center",
                 "line_rule": "exact", "color": ""},
    "abstract_body": {"size": 12.0, "font": "宋体", "ascii": "Times New Roman",
                      "bold": False, "line": 22.0, "align": "both",
                      "line_rule": "exact", "color": ""},
    "keywords": {"size": 12.0, "font": "宋体", "ascii": "Times New Roman",
                 "bold": False, "line": 22.0, "align": "both",
                 "line_rule": "exact", "color": ""},
    # 作者 / 单位 / 通讯（题目下面的信息行）。旧版没有这个角色，
    # 这些行完全不被格式化 → 行距挤在一起
    "meta": {"size": 12.0, "font": "宋体", "ascii": "Times New Roman",
             "bold": False, "line": 20.0, "align": "center",
             "line_rule": "exact", "color": ""},
    "toc": {"size": 15.0, "font": "宋体", "ascii": "Times New Roman",
            "bold": True, "line": 22.0, "align": "center",
            "line_rule": "exact", "color": ""},
    "reference": {"size": 10.5, "font": "宋体", "ascii": "Times New Roman",
                  "bold": False, "line": 18.0, "align": "both",
                  "line_rule": "exact", "color": ""},
    # 尾注区（References / Funding / Declarations / Appendix / 致谢 / 附录）：
    # 不是章节，不能占章号，否则「参考文献」会被编成第 N 章
    "back": {"size": 10.5, "font": "宋体", "ascii": "Times New Roman",
             "bold": False, "line": 18.0, "align": "both",
             "line_rule": "exact", "color": ""},
    # 公式必须用"单倍/自动"行距：固定磅值会把 19 磅高的 OLE 公式压扁
    # （实测 Word 里公式上下标挤在一起、还和编号贴在一起）
    "equation": {"size": 12.0, "font": "宋体", "ascii": "Times New Roman",
                 "bold": False, "line": 15.0, "align": "center",
                 "line_rule": "single", "color": ""},
}

# 行距规则：卡片里给用户选，不再只支持"固定值"
LINE_RULES = [("exact", "固定值（磅）"), ("single", "单倍行距"),
              ("multiple", "1.5 倍行距"), ("multiple2", "2 倍行距"),
              ("minimum", "最小值（磅）")]
# 字色候选（空 = 不指定，用 Word 的自动/黑色）
COLOR_CHOICES = [("", "自动（黑）"), ("000000", "黑色"), ("333333", "深灰"),
                 ("C00000", "深红"), ("1F4E79", "深蓝"), ("375623", "深绿")]

# Ribbon 分组定义：(组名, [(按钮 key, 显示名)])
# 每组 3~4 个，标题一律固定在卡片底部（见 main_window._ribbon）
RIBBON_GROUPS = [
    ("版面", [("page", "页面设置"), ("indent", "正文缩进"),
             ("not_indent", "段落顶格"), ("meta", "作者信息")]),
    ("标题", [("title", "论文题目"), ("heading1", "一级标题"),
              ("heading2", "二级标题"), ("heading3", "三级标题")]),
    ("正文", [("body", "正文格式"), ("equation", "公式"),
              ("figure_caption", "图注格式"), ("table_caption", "表注格式")]),
    ("图表与编号", [("table_style", "表格样式"), ("num_heading", "标题编号"),
                   ("num_figure", "图号"), ("num_table", "表号")]),
    ("页面", [("layout_section", "分节"), ("layout_pagenum", "页码"),
              ("layout_header", "页眉")]),
    ("前置部分", [("abstract", "摘要标题"), ("abstract_body", "摘要正文"),
                 ("keywords", "关键词"), ("toc", "目录")]),
    ("方案", [("schemes", "格式方案"), ("roles", "识别明细")]),
    ("其他", [("reference", "参考文献"), ("check", "格式检查"),
             ("about", "关于")]),
]

# 期刊论文：没有目录、不分节。目录卡片仍然显示但置灰（用户要求"能看到但点不了"）。
RIBBON_GROUPS_JOURNAL = [
    ("版面", [("page", "页面设置"), ("indent", "正文缩进"),
             ("not_indent", "段落顶格"), ("meta", "作者信息")]),
    ("标题", [("title", "论文题目"), ("heading1", "一级标题"),
              ("heading2", "二级标题"), ("heading3", "三级标题")]),
    ("正文", [("body", "正文格式"), ("equation", "公式"),
              ("figure_caption", "图注格式"), ("table_caption", "表注格式")]),
    ("图表与编号", [("table_style", "表格样式"), ("num_heading", "标题编号"),
                   ("num_figure", "图号"), ("num_table", "表号")]),
    ("页面", [("layout_pagenum", "页码"), ("layout_header", "页眉")]),
    ("前置部分", [("abstract", "摘要"), ("keywords", "关键词"),
                 ("toc", "目录")]),
    ("方案", [("schemes", "格式方案"), ("roles", "识别明细")]),
    ("其他", [("reference", "参考文献"), ("check", "格式检查"),
             ("about", "关于")]),
]

# 期刊形态下显示但不可点的按钮
JOURNAL_DISABLED = {"toc": "期刊论文不设目录"}
# 期刊形态下整组不显示的部分已在 RIBBON_GROUPS_JOURNAL 里去掉（如「分节」）


def ribbon_groups(profile: str = "thesis"):
    return RIBBON_GROUPS_JOURNAL if profile == "journal" else RIBBON_GROUPS

# 可以在设置卡片里改的角色
EDITABLE_ROLES = ["title", "heading1", "heading2", "heading3", "body",
                  "abstract", "abstract_body", "keywords", "meta", "toc",
                  "reference", "equation",
                  "figure_caption", "table_caption", "heading"]

ROLE_LABELS = {
    "title": "论文题目",
    "abstract": "摘要标题", "abstract_body": "摘要正文",
    "keywords": "关键词", "toc": "目录标题", "toc_item": "目录条目",
    "reference": "参考文献",
    "meta": "作者信息",
    "back": "尾注（致谢/基金/声明/附录）",
    "equation": "公式",
    "heading1": "一级标题", "heading2": "二级标题", "heading3": "三级标题",
    "body": "正文", "abstract_body": "摘要正文",
    "figure_caption": "图注", "table_caption": "表注",
    "heading": "标题（未定级）",
}

# ------------------------------------------------------------------ #
# 两种论文形态
# ------------------------------------------------------------------ #
#
# thesis  学位论文：三节（封面 / 前置 / 正文），封面无页眉页码，
#         前置罗马数字，正文阿拉伯从 1
# journal 期刊论文：单节，页码从头就是阿拉伯，不切封面/前置，默认无页眉
PROFILE_LABELS = {"thesis": "学位论文", "journal": "期刊论文"}

JOURNAL_PAGE = {"top": 25.0, "bottom": 25.0, "left": 20.0, "right": 20.0,
                "header": 15.0, "footer": 20.0}

JOURNAL_ROLES = {k: dict(v) for k, v in ROLE_DEFAULTS.items()}
JOURNAL_ROLES.update({
    "title": {"size": 16.0, "font": "宋体", "ascii": "Times New Roman",
              "bold": True, "line": 24.0, "align": "center"},
    "heading1": {"size": 14.0, "font": "宋体", "ascii": "Times New Roman",
                 "bold": True, "line": 20.0, "align": "left"},
    "heading2": {"size": 12.0, "font": "宋体", "ascii": "Times New Roman",
                 "bold": True, "line": 20.0, "align": "left"},
    "heading3": {"size": 10.5, "font": "宋体", "ascii": "Times New Roman",
                 "bold": True, "line": 20.0, "align": "left"},
    "body": {"size": 10.5, "font": "宋体", "ascii": "Times New Roman",
             "bold": False, "line": 20.0, "align": "both"},
    "abstract_body": {"size": 10.5, "font": "宋体", "ascii": "Times New Roman",
                      "bold": False, "line": 20.0, "align": "both",
                      "line_rule": "exact", "color": ""},
    "figure_caption": {"size": 9.0, "font": "宋体", "ascii": "Times New Roman",
                       "bold": True, "line": 18.0, "align": "center",
                       "line_rule": "exact", "color": ""},
    "table_caption": {"size": 9.0, "font": "宋体", "ascii": "Times New Roman",
                      "bold": True, "line": 18.0, "align": "center",
                      "line_rule": "exact", "color": ""},
    "heading": {"size": 14.0, "font": "宋体", "ascii": "Times New Roman",
                "bold": True, "line": 20.0, "align": "left"},
    # 期刊的「摘要 / 关键词」是段落开头的一行，不是居中标题 → 居左
    "abstract": {"size": 10.5, "font": "宋体", "ascii": "Times New Roman",
                 "bold": False, "line": 20.0, "align": "left",
                 "line_rule": "exact", "color": ""},
    "keywords": {"size": 10.5, "font": "宋体", "ascii": "Times New Roman",
                 "bold": False, "line": 20.0, "align": "left",
                 "line_rule": "exact", "color": ""},
    "meta": {"size": 10.5, "font": "宋体", "ascii": "Times New Roman",
             "bold": False, "line": 18.0, "align": "center",
             "line_rule": "exact", "color": ""},
    # 期刊没有目录，但卡片仍然显示（置灰），所以角色要保留
    "toc": {"size": 15.0, "font": "宋体", "ascii": "Times New Roman",
            "bold": True, "line": 22.0, "align": "center",
            "line_rule": "exact", "color": ""},
})

# 切换 profile 时一起换掉的项（其余项保持用户当前的值）
PROFILE_KEYS = ("page", "roles", "header_text", "enable_header",
                "enable_sections", "enable_page_numbers", "first_line_indent_chars",
                "bold_label")

# 补上新字段（老配置里没有）
for _roles in (ROLE_DEFAULTS, JOURNAL_ROLES):
    for _v in _roles.values():
        _v.setdefault("line_rule", "exact")
        _v.setdefault("color", "")


@dataclass
class Settings:
    page: dict = field(default_factory=lambda: dict(PAGE_DEFAULT))
    roles: dict = field(default_factory=lambda: {k: dict(v) for k, v in ROLE_DEFAULTS.items()})
    paper: str = PAPER
    line_spacing: float = LINE_SPACING_PT
    header_text: str = HEADER_TEXT
    enable_header: bool = True      # 关掉就没有页眉（封面本来也没有）
    # 页眉横线：0 = 没有；否则是线宽（磅）
    header_line: float = 0.75
    header_line_color: str = "000000"
    header_line_style: str = "single"   # single/double/dashed/dotted
    # 页眉文字的字体 / 字号 / 与横线的间距（磅）
    header_font: str = "宋体"
    header_ascii: str = "Times New Roman"
    header_size: float = 10.5
    header_bold: bool = False
    header_color: str = ""              # 空 = 自动（黑）
    header_gap: float = 4.0             # 页眉文字与横线之间的距离
    # 表格：three_line 三线表（两粗一细：顶/底 1.5 磅、栏目线 0.75 磅）/ grid 网格 / none 无线
    table_style: str = "three_line"
    table_line: float = 1.5
    table_line_inner: float = 0.75
    table_font_size: float = 10.5
    table_font: str = "宋体"
    table_ascii: str = "Times New Roman"
    table_header_bold: bool = True
    table_bold: bool = False
    table_color: str = ""
    # 表内文字对齐 / 表格在页面里的位置 / 清除原表底纹（去掉表头蓝底等填充色）
    table_align: str = "left"
    table_pos: str = "center"
    table_clear_shading: bool = True

    # 图号 / 表号自定义
    caption_fig_prefix: str = "图"
    caption_tab_prefix: str = "表"
    # 原稿图注普遍用 Fig./Table 开头时，自动沿用英文前缀（取消勾选则用上面设的）
    caption_prefix_auto: bool = True
    caption_sep: str = "."          # "." / "-" / ""（图1-1、图11）
    caption_by_chapter: bool = True
    caption_start: int = 1
    # 标题编号里"章号与节号"的分隔（1.1 / 1-1 / 11）
    heading_sep: str = "."
    # 段落开头 Abstract: / 关键词: 这种标签是否加粗（期刊论文习惯加粗）
    bold_label: bool = False
    # 自动把 CO2 / m2 这类纯文本里的数字转成上/下标
    enable_marks: bool = True
    # 开关
    enable_sections: bool = True
    enable_page_numbers: bool = True
    # 编号默认关闭：只有用户明确勾选时才补号/重排，避免改坏原稿已有编号
    enable_number_heading: bool = False
    enable_number_figure: bool = False
    enable_number_table: bool = False
    # 原稿里已经写着的编号：保留并补缺（默认），还是全部替换重排
    number_keep_existing: bool = True
    # 正文缩进（首行缩进字符数）
    first_line_indent_chars: float = 2.0
    apply_indent: bool = True
    # 段落顶格（取消首行缩进）
    apply_not_indent: bool = False
    # 论文形态：thesis 学位论文 / journal 期刊论文
    profile: str = "journal"
    # 各 profile 的取值快照，切换时互相换入换出
    _profiles: dict = field(default_factory=dict)

    # ---------------------------------------------------------------- #
    def __post_init__(self) -> None:
        """默认形态是期刊论文，字段初值还是学位论文那套，这里要换成期刊预设。"""
        if self.profile == "journal":
            self._load_profile_preset("journal")

    def snapshot(self) -> dict:
        return {k: (dict(v) if isinstance(v, dict) else v)
                for k, v in ((k, getattr(self, k)) for k in PROFILE_KEYS)}

    def set_profile(self, profile: str) -> bool:
        """切换论文形态。当前取值存回旧 profile，再载入新 profile 的取值。"""
        if profile not in PROFILE_LABELS or profile == self.profile:
            return False
        self._profiles[self.profile] = self.snapshot()
        saved = self._profiles.get(profile)
        if saved is None:
            self._load_profile_preset(profile)
        else:
            for k, v in saved.items():
                setattr(self, k, dict(v) if isinstance(v, dict) else v)
        self.profile = profile
        return True

    def _load_profile_preset(self, profile: str) -> None:
        if profile == "journal":
            self.page = dict(JOURNAL_PAGE)
            self.roles = {k: dict(v) for k, v in JOURNAL_ROLES.items()}
            self.header_text = ""
            self.enable_header = False
            self.enable_sections = False
            self.enable_page_numbers = True
            self.first_line_indent_chars = 2.0
            self.bold_label = True
        else:
            self.page = dict(PAGE_DEFAULT)
            self.roles = {k: dict(v) for k, v in ROLE_DEFAULTS.items()}
            self.header_text = HEADER_TEXT
            self.enable_header = True
            self.enable_sections = True
            self.enable_page_numbers = True
            self.first_line_indent_chars = 2.0
            self.bold_label = False

    def to_json(self) -> dict:
        return asdict(self)

    def scheme_payload(self) -> dict:
        """存成"格式方案"的完整设置（去掉内存里的 profile 快照）。"""
        data = self.to_json()
        data.pop("_profiles", None)
        return data

    @classmethod
    def from_json(cls, data: dict) -> "Settings":
        s = cls()
        if "page" in data:
            s.page.update(data["page"])
        for k, v in (data.get("roles") or {}).items():
            if k in s.roles:
                s.roles[k].update(v)
        for key in ("paper", "line_spacing", "header_text", "enable_sections",
                    "enable_page_numbers", "enable_header", "profile",
                    "enable_number_heading",
                    "enable_number_figure", "enable_number_table",
                    "number_keep_existing",
                    "first_line_indent_chars", "apply_indent", "apply_not_indent",
                    "header_line", "header_line_color", "header_line_style",
                    "header_font", "header_ascii", "header_size", "header_bold",
                    "header_color", "header_gap",
                    "table_style", "table_line", "table_line_inner", "table_font_size",
                    "table_font", "table_ascii", "table_header_bold", "table_bold",
                    "table_color", "table_align", "table_pos",
                    "caption_fig_prefix", "caption_tab_prefix", "caption_sep",
                    "caption_by_chapter", "caption_start", "heading_sep",
                    "enable_marks", "bold_label", "caption_prefix_auto"):
            if key in data:
                setattr(s, key, data[key])
        # 兼容老版本方案：没有单独的 table_header_bold 时沿用原 table_bold
        if "table_header_bold" not in data and "table_bold" in data:
            s.table_header_bold = bool(data["table_bold"])
        return s


DEFAULT_SETTINGS = Settings()

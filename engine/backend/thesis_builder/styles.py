"""所有格式常量集中管理，方便前端配置覆盖。

山西财经大学学年论文格式规范：
- 纸张 A4，纵向，左侧装订
- 页边距：上3cm 下2.5cm 左2.5cm 右2cm
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Dict, Any


# ═══════════════════════════════════════════
# 字号映射（Word 半磅 → python-docx Pt）
# ═══════════════════════════════════════════
FONT_SIZE_MAP = {
    "小二号": 18,
    "三号": 16,
    "四号": 14,
    "小四号": 12,
    "五号": 10.5,
    "小五号": 9,
}


# ═══════════════════════════════════════════
# 字体名称常量
# ═══════════════════════════════════════════
FONT_SIMSUN = "SimSun"           # 宋体
FONT_SIMHEI = "SimHei"           # 黑体
FONT_FANGSONG = "FangSong"       # 仿宋
FONT_TIMES = "Times New Roman"   # 西文默认


# ═══════════════════════════════════════════
# 页面设置
# ═══════════════════════════════════════════
@dataclass
class PageConfig:
    """页面尺寸与页边距（单位：厘米）"""
    width_cm: float = 21.0
    height_cm: float = 29.7
    margin_top_cm: float = 3.0
    margin_bottom_cm: float = 2.5
    margin_left_cm: float = 2.5
    margin_right_cm: float = 2.0


# ═══════════════════════════════════════════
# 字体配置
# ═══════════════════════════════════════════
@dataclass
class FontConfig:
    """所有字体相关配置"""
    # 正文
    body_east_asian: str = FONT_SIMSUN
    body_latin: str = FONT_TIMES
    body_size_pt: float = 12          # 小四号

    # 一级标题
    heading1_font: str = FONT_SIMHEI
    heading1_size_pt: float = 16      # 三号
    heading1_bold: bool = False       # 标准文档黑体不加粗

    # 二级标题
    heading2_font: str = FONT_SIMSUN
    heading2_size_pt: float = 14      # 四号
    heading2_bold: bool = True

    # 三级标题
    heading3_font: str = FONT_SIMSUN
    heading3_size_pt: float = 12      # 小四号
    heading3_bold: bool = True

    # 摘要标题
    abstract_title_font: str = FONT_SIMSUN  # 规范：宋体小二加粗
    abstract_title_size_pt: float = 18  # 小二号

    # 英文摘要标题
    eng_abstract_title_font: str = FONT_TIMES
    eng_abstract_title_size_pt: float = 16  # 三号

    # 关键词标签
    keywords_label_font: str = FONT_SIMHEI
    keywords_label_size_pt: float = 14  # 四号

    # 参考文献
    ref_font: str = FONT_SIMSUN
    ref_size_pt: float = 12            # 小四号

    # 图表标题
    caption_font: str = FONT_FANGSONG
    caption_size_pt: float = 10.5      # 五号

    # 页眉
    header_font_east_asian: str = FONT_SIMSUN
    header_font_latin: str = FONT_TIMES
    header_size_pt: float = 9          # 小五号

    # 页脚（页码）
    footer_font_east_asian: str = FONT_SIMSUN
    footer_font_latin: str = FONT_TIMES
    footer_size_pt: float = 9          # 小五号

    # 封面字段
    cover_field_font: str = FONT_SIMSUN
    cover_field_size_pt: float = 14    # 四号
    cover_field_bold: bool = True


# ═══════════════════════════════════════════
# 段落间距配置
# ═══════════════════════════════════════════
@dataclass
class SpacingConfig:
    """段落间距与行距"""
    line_spacing: float = 1.25         # 1.25倍行距

    # 一级标题段前段后（1行 ≈ 18pt at 小四号正文）
    heading1_before_pt: float = 18
    heading1_after_pt: float = 18

    # 二级标题
    heading2_before_pt: float = 0
    heading2_after_pt: float = 0

    # 三级标题
    heading3_before_pt: float = 0
    heading3_after_pt: float = 0

    # 摘要标题段前段后
    abstract_before_pt: float = 18
    abstract_after_pt: float = 18

    # 关键词段前
    keywords_before_pt: float = 12     # 与标准文档一致

    # 首行缩进字符数
    first_line_indent_chars: int = 2

    # 参考文献悬挂缩进
    ref_hanging_indent_chars: int = 2

    # 图表标题段后
    caption_after_pt: float = 6        # 0.5行


# ═══════════════════════════════════════════
# 页眉页脚配置
# ═══════════════════════════════════════════
@dataclass
class HeaderConfig:
    """页眉设置"""
    text_template: str = "山西财经大学{grade}级本科生学年论文"
    odd_align: str = "right"
    even_align: str = "left"
    show: bool = True


@dataclass
class FooterConfig:
    """页脚（页码）设置"""
    show_page_number: bool = True
    abstract_format: str = "roman_upper"   # 大写罗马数字
    body_format: str = "decimal"           # 阿拉伯数字
    abstract_align: str = "center"
    odd_align: str = "right"
    even_align: str = "left"


# ═══════════════════════════════════════════
# 章节开关
# ═══════════════════════════════════════════
@dataclass
class SectionsConfig:
    """控制是否生成各章节"""
    cover: bool = True
    abstract_cn: bool = True
    abstract_en: bool = True
    toc: bool = True
    body: bool = True
    references: bool = True
    appendix: bool = True
    acknowledgment: bool = True
    back_cover: bool = True


# ═══════════════════════════════════════════
# 汇总配置
# ═══════════════════════════════════════════
@dataclass
class ThesisConfig:
    """完整排版配置"""
    page: PageConfig = field(default_factory=PageConfig)
    fonts: FontConfig = field(default_factory=FontConfig)
    spacing: SpacingConfig = field(default_factory=SpacingConfig)
    header: HeaderConfig = field(default_factory=HeaderConfig)
    footer: FooterConfig = field(default_factory=FooterConfig)
    sections: SectionsConfig = field(default_factory=SectionsConfig)

    @classmethod
    def from_dict(cls, d: Dict[str, Any]) -> "ThesisConfig":
        """从字典创建配置（支持前端传入的部分配置）"""
        cfg = cls()
        if "page" in d:
            for k, v in d["page"].items():
                if hasattr(cfg.page, k):
                    setattr(cfg.page, k, v)
        if "fonts" in d:
            for k, v in d["fonts"].items():
                if hasattr(cfg.fonts, k):
                    setattr(cfg.fonts, k, v)
        if "spacing" in d:
            for k, v in d["spacing"].items():
                if hasattr(cfg.spacing, k):
                    setattr(cfg.spacing, k, v)
        if "header" in d:
            for k, v in d["header"].items():
                if hasattr(cfg.header, k):
                    setattr(cfg.header, k, v)
        if "footer" in d:
            for k, v in d["footer"].items():
                if hasattr(cfg.footer, k):
                    setattr(cfg.footer, k, v)
        if "sections" in d:
            for k, v in d["sections"].items():
                if hasattr(cfg.sections, k):
                    setattr(cfg.sections, k, v)
        return cfg

    def to_dict(self) -> Dict[str, Any]:
        """导出为字典"""
        from dataclasses import asdict
        return asdict(self)

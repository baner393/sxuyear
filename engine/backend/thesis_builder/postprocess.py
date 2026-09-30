"""后处理：页眉页脚、分节符、脚注、批注、目录域。

对应 C# 中的 MakeHeader / MakeFooter / CloseSection / AddTOC /
SetupFootnotes / AddFnContent / AddComments 等函数。
"""

from __future__ import annotations

import re
from pathlib import Path
from typing import List, Optional

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml
from lxml import etree

from .styles import (
    ThesisConfig, FONT_SIMSUN, FONT_TIMES, FONT_SIMHEI,
)


# ═══════════════════════════════════════════
# 页眉页脚
# ═══════════════════════════════════════════

def setup_headers_footers(
    doc: Document, cfg: ThesisConfig, grade: str,
    section_types: Optional[List[str]] = None,
    title: Optional[str] = None,
) -> None:
    """为所有分节设置页眉和页脚，支持奇偶页不同对齐。

    Section 类型说明：
    - "cover"        : 封面/封底，无页眉无页脚无页码
    - "front_matter" : 中文摘要、英文摘要、目录，页码罗马数字居中
    - "body"         : 正文、参考文献、附录、致谢，页码阿拉伯数字奇右偶左

    Args:
        doc: 输出文档
        cfg: 排版配置
        grade: 年级（如 "2023"，取自学号前 4 位＝入学年）
        section_types: 每个分节的类型列表，与 doc.sections 一一对应
        title: 论文题目（页眉模板可用 {title}，川大完整版/上二工大等
               校规范页眉含论文题目）
    """
    # 页眉模板可用三个变量：
    #   {grade}     入学年（学年论文用「2023级」）
    #   {grad_year} 毕业届＝入学年+4（毕业论文规范要求「2027届」）
    #   {title}     论文题目
    try:
        grad_year = str(int(grade) + 4)
    except (TypeError, ValueError):
        grad_year = grade
    try:
        header_text = cfg.header.text_template.format(
            grade=grade, grad_year=grad_year, title=title or '')
    except (KeyError, IndexError, ValueError):
        # 模板档案里写了不认识的占位符时按原样输出，不让转换失败
        header_text = cfg.header.text_template

    # 文档级设置：奇偶页不同页眉页脚 + 行尾下划线空格渲染
    _setup_document_settings(doc)

    sections = list(doc.sections)

    if section_types is None:
        section_types = ["body"] * len(sections)

    # 确保长度匹配
    while len(section_types) < len(sections):
        section_types.append("body")

    # 页码分组连续：罗马页码组（前置部分）和阿拉伯页码组（正文起）
    # 各自只在组内第一个分节 start=1，后续分节沿用上一节页码继续编号
    front_started = False
    body_started = False

    for i, section in enumerate(sections):
        sec_type = section_types[i] if i < len(section_types) else "body"

        # 脚注：每页重新从 1 编号（规范第10条）
        _set_footnote_props(section)

        if sec_type == "cover":
            _setup_cover_section(section)
        elif sec_type == "front_matter":
            _setup_front_matter_section(
                section, header_text, cfg,
                restart=not front_started,
            )
            front_started = True
        else:  # "body"
            _setup_body_section(
                section, header_text, cfg,
                restart=not body_started,
            )
            body_started = True


def _setup_document_settings(doc: Document) -> None:
    """写入文档级 settings.xml 设置（与学校标准文档一致）。

    - w:evenAndOddHeaders：启用奇偶页不同页眉页脚的总开关，
      没有它，分节里写好的偶数页页眉页脚不会生效。
    - w:compat/w:ulTrailSpace：渲染行尾的下划线空格。
      承诺/授权页的签名线是「带下划线的空格」，行尾部分没有此开关不显示。

    注意：必须改 doc.settings.element（保存时按 element 序列化），
    直接写 settings part 的 _blob 会被丢弃。
    """
    settings_el = doc.settings.element

    # compat 元素（无则创建，追加到末尾——Word 对顺序宽容）
    compat = settings_el.find(qn("w:compat"))
    if compat is None:
        compat = parse_xml(f'<w:compat {nsdecls("w")}/>')
        settings_el.append(compat)

    # 奇偶页不同（插到 compat 之前，接近 schema 顺序）
    if settings_el.find(qn("w:evenAndOddHeaders")) is None:
        even_odd = parse_xml(f'<w:evenAndOddHeaders {nsdecls("w")}/>')
        compat.addprevious(even_odd)

    # 行尾下划线空格（compat 内子元素，需在 compatSetting 之前）
    if compat.find(qn("w:ulTrailSpace")) is None:
        ul = parse_xml(f'<w:ulTrailSpace {nsdecls("w")}/>')
        compat.insert(0, ul)


def _set_footnote_props(section) -> None:
    """设置分节脚注属性：每页重新从 1 开始编号。"""
    sectPr = section._sectPr
    if sectPr.find(qn("w:footnotePr")) is None:
        fp = parse_xml(
            f'<w:footnotePr {nsdecls("w")}>'
            '<w:numRestart w:val="eachPage"/>'
            '</w:footnotePr>'
        )
        sectPr.insert(0, fp)


def _setup_cover_section(section) -> None:
    """封面分节：清空奇偶页的页眉和页脚，不显示任何页码。"""
    for attr in ("header", "even_page_header"):
        h = getattr(section, attr)
        h.is_linked_to_previous = False
        for p in h.paragraphs:
            p.clear()

    for attr in ("footer", "even_page_footer"):
        f = getattr(section, attr)
        f.is_linked_to_previous = False
        for p in f.paragraphs:
            p.clear()


def _setup_front_matter_section(
    section, header_text: str, cfg: ThesisConfig, restart: bool = False,
) -> None:
    """摘要 / 目录分节：页眉奇右偶左，页脚罗马数字居中。"""
    _set_pg_num_type(section, "upperRoman", 1 if restart else None)

    # 页眉（奇数页右对齐 / 偶数页左对齐）
    _set_header_content(
        section.header, header_text, cfg, WD_ALIGN_PARAGRAPH.RIGHT,
    )
    _set_header_content(
        section.even_page_header, header_text, cfg, WD_ALIGN_PARAGRAPH.LEFT,
    )

    # 页脚（居中）
    _set_footer_page_number(
        section.footer, cfg, WD_ALIGN_PARAGRAPH.CENTER,
    )
    _set_footer_page_number(
        section.even_page_footer, cfg, WD_ALIGN_PARAGRAPH.CENTER,
    )


def _setup_body_section(
    section, header_text: str, cfg: ThesisConfig, restart: bool = False,
) -> None:
    """正文 / 参考文献 / 附录 / 致谢分节：页眉奇右偶左，页码阿拉伯数字奇右偶左。"""
    _set_pg_num_type(section, "decimal", 1 if restart else None)

    # 页眉（奇数页右对齐 / 偶数页左对齐）
    _set_header_content(
        section.header, header_text, cfg, WD_ALIGN_PARAGRAPH.RIGHT,
    )
    _set_header_content(
        section.even_page_header, header_text, cfg, WD_ALIGN_PARAGRAPH.LEFT,
    )

    # 页脚（奇数页右 / 偶数页左）
    _set_footer_page_number(
        section.footer, cfg, WD_ALIGN_PARAGRAPH.RIGHT,
    )
    _set_footer_page_number(
        section.even_page_footer, cfg, WD_ALIGN_PARAGRAPH.LEFT,
    )


def _set_pg_num_type(section, fmt: str, start: Optional[int] = None) -> None:
    """设置分节的页码格式（upperRoman / decimal）。

    start 为 None 时不写 w:start，页码沿用上一分节继续编号。
    """
    sectPr = section._sectPr
    for old in sectPr.findall(qn("w:pgNumType")):
        sectPr.remove(old)
    start_attr = f' w:start="{start}"' if start is not None else ""
    pg_num = parse_xml(
        f'<w:pgNumType {nsdecls("w")} w:fmt="{fmt}"{start_attr}/>'
    )
    sectPr.append(pg_num)


def _set_header_content(header, text: str, cfg: ThesisConfig, alignment) -> None:
    """设置单个页眉对象的文本和样式（含底部边框线）。"""
    header.is_linked_to_previous = False
    for p in header.paragraphs:
        p.clear()

    para = header.paragraphs[0]
    para.alignment = alignment

    run = para.add_run(text)
    run.font.size = Pt(cfg.fonts.header_size_pt)
    run.font.name = cfg.fonts.header_font_latin
    rpr = run._element.get_or_add_rPr()
    rFonts = rpr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = parse_xml(f'<w:rFonts {nsdecls("w")}/>')
        rpr.insert(0, rFonts)
    rFonts.set(qn("w:eastAsia"), cfg.fonts.header_font_east_asian)

    # 底部边框线
    pPr = para._element.get_or_add_pPr()
    pBdr = parse_xml(
        f'<w:pBdr {nsdecls("w")}>'
        f'  <w:bottom w:val="single" w:sz="6" w:space="1" w:color="000000"/>'
        f'</w:pBdr>'
    )
    pPr.append(pBdr)


def _set_footer_page_number(footer, cfg: ThesisConfig, alignment) -> None:
    """设置单个页脚对象的 PAGE 域代码。"""
    footer.is_linked_to_previous = False
    for p in footer.paragraphs:
        p.clear()

    para = footer.paragraphs[0]
    para.alignment = alignment

    # 使用 XML 直接插入 PAGE 域代码
    fld_char_begin = parse_xml(
        f'<w:fldChar {nsdecls("w")} w:fldCharType="begin"/>'
    )
    fld_code = parse_xml(
        f'<w:instrText {nsdecls("w")} xml:space="preserve"> PAGE </w:instrText>'
    )
    fld_char_end = parse_xml(
        f'<w:fldChar {nsdecls("w")} w:fldCharType="end"/>'
    )

    half_pt = str(int(cfg.fonts.footer_size_pt * 2))

    run1 = para._element.makeelement(qn("w:r"), {})
    rPr1 = run1.makeelement(qn("w:rPr"), {})
    sz1 = rPr1.makeelement(qn("w:sz"), {qn("w:val"): half_pt})
    rPr1.append(sz1)
    run1.append(rPr1)
    run1.append(fld_char_begin)
    para._element.append(run1)

    run2 = para._element.makeelement(qn("w:r"), {})
    rPr2 = run2.makeelement(qn("w:rPr"), {})
    sz2 = rPr2.makeelement(qn("w:sz"), {qn("w:val"): half_pt})
    rPr2.append(sz2)
    run2.append(rPr2)
    run2.append(fld_code)
    para._element.append(run2)

    run3 = para._element.makeelement(qn("w:r"), {})
    rPr3 = run3.makeelement(qn("w:rPr"), {})
    sz3 = rPr3.makeelement(qn("w:sz"), {qn("w:val"): half_pt})
    rPr3.append(sz3)
    run3.append(rPr3)
    run3.append(fld_char_end)
    para._element.append(run3)


# ═══════════════════════════════════════════
# 分节符
# ═══════════════════════════════════════════

def add_section_break(doc: Document, break_type: str = "new_page") -> None:
    """在文档末尾添加分节符。

    Args:
        doc: 文档对象
        break_type: 分节类型，new_page / continuous / even_page / odd_page

    Raises:
        ValueError: break_type 不是已知类型。以前未知值静默回退 nextPage，
            拼写错误会无声退化成「下一页」且无任何校验兜底，故改为报错。
    """
    new_section = doc.add_section()
    sectPr = new_section._sectPr

    # 直接操作 XML 设置分节符类型，避免 python-docx 版本间枚举值差异
    _TYPE_MAP = {
        "new_page": "nextPage",
        "continuous": "continuous",
        "even_page": "evenPage",
        "odd_page": "oddPage",
    }
    if break_type not in _TYPE_MAP:
        raise ValueError(
            f"未知分节类型 {break_type!r}，可选: {', '.join(_TYPE_MAP)}"
        )
    val = _TYPE_MAP[break_type]

    # 查找已有的 w:type 元素并修改 w:val 属性
    type_elem = sectPr.find(qn("w:type"))
    if type_elem is not None:
        type_elem.set(qn("w:val"), val)
    else:
        # ⚠️ 必须按 CT_SectPr 的 schema 顺序插入（headerReference/footerReference/
        # footnotePr/endnotePr 之后、pgSz 之前）。曾用 insert(0) 把 w:type 插到
        # footnotePr 前面，2026-07-20 实测该顺序违规会让 Word 的 ExportAsFixedFormat
        # 陷入死循环（>4 分钟满核占用，正常 11 秒），且 Word 打开文档时无任何报错。
        type_elem = parse_xml(
            f'<w:type {nsdecls("w")} w:val="{val}"/>'
        )
        sectPr.insert_element_before(
            type_elem,
            "w:pgSz", "w:pgMar", "w:paperSrc", "w:pgBorders", "w:lnNumType",
            "w:pgNumType", "w:cols", "w:formProt", "w:vAlign", "w:noEndnote",
            "w:titlePg", "w:textDirection", "w:bidi", "w:rtlGutter",
            "w:docGrid", "w:printerSettings", "w:sectPrChange",
        )


# ═══════════════════════════════════════════
# 脚注
# ═══════════════════════════════════════════

def add_footnotes(doc: Document, paragraph, footnote_texts: List[str],
                  start_id: int = 1) -> int:
    """为段落添加脚注。

    Args:
        doc: 文档对象
        paragraph: 要添加脚注的段落
        footnote_texts: 脚注内容列表
        start_id: 起始脚注 ID

    Returns:
        下一个可用的脚注 ID
    """
    # 获取或创建 footnotes part
    footnotes_part = _get_or_create_footnotes_part(doc)
    fn_id = start_id

    for fn_text in footnote_texts:
        # 在段落中添加脚注引用（上标）
        run = paragraph.add_run()
        run.font.superscript = True
        run.font.size = Pt(9)  # 小五号

        # 插入 footnoteReference XML
        fn_ref = parse_xml(
            f'<w:footnoteReference {nsdecls("w")} w:id="{fn_id}"/>'
        )
        run._element.append(fn_ref)

        # 添加脚注内容
        _add_footnote_content(footnotes_part, fn_id, fn_text)
        fn_id += 1

    return fn_id


def _get_or_create_footnotes_part(doc: Document):
    """获取或创建 footnotes part"""
    from docx.opc.part import Part
    from docx.opc.constants import RELATIONSHIP_TYPE as RT

    # 检查是否已存在
    for rel in doc.part.rels.values():
        if 'footnotes' in rel.reltype:
            return rel.target_part

    # 创建新的 footnotes part
    footnotes_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '  <w:footnote w:type="separator" w:id="-1">'
        '    <w:p><w:r><w:separator/></w:r></w:p>'
        '  </w:footnote>'
        '  <w:footnote w:type="continuationSeparator" w:id="0">'
        '    <w:p><w:r><w:continuationSeparator/></w:r></w:p>'
        '  </w:footnote>'
        '</w:footnotes>'
    )

    from docx.opc.part import Part
    from docx.opc.packuri import PackURI
    part = Part(
        PackURI('/word/footnotes.xml'),
        'application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml',
        footnotes_xml.encode('utf-8'),
        doc.part.package,
    )
    doc.part.relate_to(part, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes')
    return part


def _add_footnote_content(footnotes_part, fn_id: int, fn_text: str) -> None:
    """向 footnotes part 添加脚注内容"""
    # 解析现有 XML
    xml_bytes = footnotes_part.blob
    root = etree.fromstring(xml_bytes)

    # 创建新脚注
    nsmap = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    fn = etree.SubElement(root, qn('w:footnote'))
    fn.set(qn('w:id'), str(fn_id))

    p = etree.SubElement(fn, qn('w:p'))
    pPr = etree.SubElement(p, qn('w:pPr'))
    spacing = etree.SubElement(pPr, qn('w:spacing'))
    spacing.set(qn('w:line'), '240')
    spacing.set(qn('w:lineRule'), 'auto')

    # 脚注引用标记（上标）
    r1 = etree.SubElement(p, qn('w:r'))
    rPr1 = etree.SubElement(r1, qn('w:rPr'))
    vAlign = etree.SubElement(rPr1, qn('w:vertAlign'))
    vAlign.set(qn('w:val'), 'superscript')
    fn_ref_mark = etree.SubElement(r1, qn('w:footnoteRef'))

    # 脚注文本
    r2 = etree.SubElement(p, qn('w:r'))
    rPr2 = etree.SubElement(r2, qn('w:rPr'))
    rFonts = etree.SubElement(rPr2, qn('w:rFonts'))
    rFonts.set(qn('w:eastAsia'), FONT_SIMSUN)
    rFonts.set(qn('w:ascii'), FONT_TIMES)
    rFonts.set(qn('w:hAnsi'), FONT_TIMES)
    sz = etree.SubElement(rPr2, qn('w:sz'))
    sz.set(qn('w:val'), '18')  # 小五号 = 9pt = 18 half-points
    t = etree.SubElement(r2, qn('w:t'))
    t.text = ' ' + fn_text
    t.set(qn('xml:space'), 'preserve')

    # 更新 part 内容
    footnotes_part._blob = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)


# ═══════════════════════════════════════════
# 目录域
# ═══════════════════════════════════════════

def add_toc_field(doc: Document) -> None:
    """添加 TOC 域代码（显示占位文本，用户在 Word 中右键更新）。

    对应 C# 中的 AddTOC 函数。
    """
    para = doc.add_paragraph()

    # 域代码：TOC \o "1-2" \h \z \u
    # Begin
    run1 = para.add_run()
    fld_char_begin = parse_xml(f'<w:fldChar {nsdecls("w")} w:fldCharType="begin"/>')
    run1._element.append(fld_char_begin)

    # Field code（注意：域代码开关是单反斜杠，如 \o "1-2"）
    run2 = para.add_run()
    instr_text = parse_xml(f'<w:instrText {nsdecls("w")} xml:space="preserve"> TOC \\o &quot;1-2&quot; \\h \\z \\u </w:instrText>')
    run2._element.append(instr_text)

    # Separate
    run3 = para.add_run()
    fld_char_sep = parse_xml(f'<w:fldChar {nsdecls("w")} w:fldCharType="separate"/>')
    run3._element.append(fld_char_sep)

    # 占位文本
    run4 = para.add_run("目录将在Word中自动生成，请右键目录→更新域")
    run4.font.size = Pt(12)

    # End
    run5 = para.add_run()
    fld_char_end = parse_xml(f'<w:fldChar {nsdecls("w")} w:fldCharType="end"/>')
    run5._element.append(fld_char_end)

"""DOCX 读取器：把 .docx 解析为中间 Block 列表。

改造自 paper_processor 的 extract.py（github.com/baner393/paper_processor），
纯 zipfile + lxml 实现，不依赖 python-docx。

职责边界：只负责「读」——文本/标题层级/表格/图片/公式/脚注/封面字段/锚点；
md 契约组装在 builder.py。
"""

from __future__ import annotations

import re
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Optional, Tuple

from lxml import etree

# ── OOXML 命名空间 ──
NSMAP = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'm': 'http://schemas.openxmlformats.org/officeDocument/2006/math',
    'v': 'urn:schemas-microsoft-com:vml',
}


def q(tag: str) -> str:
    """快捷命名空间查询，如 q('w:p') → {ns}p"""
    ns, name = tag.split(':')
    return f'{{{NSMAP[ns]}}}{name}'


# 封面表格标签 → 元数据字段（标签先去掉全部空白再匹配；顺序即优先级，
# 「专业代码」必须先于「专业」）
_COVER_LABEL_MAP = [
    ('中文题目', 'title'),
    ('论文题目', 'title'),
    ('英文题目', 'engTitle'),
    ('专业代码', 'majorCode'),
    ('学生姓名', 'name'),
    ('姓名', 'name'),
    ('学号', 'studentId'),
    ('班级', 'class'),
    ('专业', 'major'),
    ('学院', 'college'),
    ('指导教师', 'advisor'),
    ('完成时间', 'date'),
    ('完成日期', 'date'),
    ('日期', 'date'),
]

# 允许导出的图片格式（converter 端 python-docx 能插入的位图）
_IMAGE_EXTS = {'.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tif', '.tiff'}


@dataclass
class Block:
    """文档中间块"""
    kind: str                 # 'heading' | 'para' | 'table' | 'image'
    level: int = 0            # heading 层级（1-3）
    text: str = ""            # 段落文本（可含 [^N] 与 <!--FORMULA:N--> 内联标记）
    table: Optional[List[List[str]]] = None
    image_rids: List[str] = field(default_factory=list)
    index: int = 0            # body 顶层元素索引


def _norm(text: str) -> str:
    """归一化：去掉全部空白（含全角空格）用于锚点/标签匹配"""
    return re.sub(r'[\s　]+', '', text)


class DocxReader:
    def __init__(self, docx_path: Path):
        self.docx_path = Path(docx_path)
        self.zf = zipfile.ZipFile(self.docx_path, 'r')
        self.body = etree.parse(self.zf.open('word/document.xml')).getroot().find(q('w:body'))
        if self.body is None:
            raise ValueError("docx 缺少 body 元素")
        self.warnings: List[str] = []
        self.image_map: Dict[str, str] = {}       # rId → zip 内路径
        self.formulas: List[bytes] = []           # 序列化的 OMML，下标+1 = N
        self.footnote_map: Dict[str, int] = {}    # w:id → 重排后的 N
        self.footnotes: Dict[int, str] = {}       # N → 文本
        self._parse_rels()
        self._parse_footnotes()

    # ── 关系与部件 ──

    def _parse_rels(self) -> None:
        """rId → 图片路径（zip 内）"""
        try:
            rels = etree.parse(self.zf.open('word/_rels/document.xml.rels'))
        except KeyError:
            return
        for rel in rels.getroot():
            if 'image' in rel.get('Type', ''):
                target = rel.get('Target', '').lstrip('/')
                if not target.startswith('word/'):
                    target = 'word/' + target
                self.image_map[rel.get('Id', '')] = target

    def _parse_footnotes(self) -> None:
        """word/footnotes.xml → 重排为 1..n"""
        try:
            root = etree.parse(self.zf.open('word/footnotes.xml')).getroot()
        except KeyError:
            return
        n = 0
        for fn in root.findall(q('w:footnote')):
            if fn.get(q('w:type')) in ('separator', 'continuationSeparator'):
                continue
            text = ' '.join(
                t for t in (self._plain_text(p).strip() for p in fn.findall(q('w:p'))) if t
            ).strip()
            if not text:
                continue
            n += 1
            self.footnote_map[fn.get(q('w:id'), '')] = n
            self.footnotes[n] = text

    # ── 文本提取 ──

    def _plain_text(self, element) -> str:
        """递归纯文本（不含公式/脚注标记），用于脚注内容、锚点判断等"""
        tag = element.tag
        if tag == q('w:t'):
            return element.text or ''
        if tag in (q('w:tab'), q('w:br')):
            return ' '
        if tag in (q('w:pPr'), q('w:rPr'), q('w:tblPr'), q('w:tcPr'),
                   q('w:trPr'), q('w:tblGrid'), q('w:sectPr')):
            return ''
        return ''.join(self._plain_text(c) for c in element)

    def _run_text(self, run) -> str:
        """run 文本；脚注引用输出 [^N] 标记"""
        parts = []
        for child in run:
            tag = child.tag
            if tag == q('w:t'):
                parts.append(child.text or '')
            elif tag in (q('w:tab'), q('w:br')):
                parts.append(' ')
            elif tag == q('w:footnoteReference'):
                fn = self.footnote_map.get(child.get(q('w:id'), ''))
                if fn:
                    parts.append(f'[^{fn}]')
        return ''.join(parts)

    def _register_formula(self, math_el) -> str:
        """序列化 OMML（带命名空间声明），返回内联占位标记"""
        self.formulas.append(etree.tostring(math_el))
        return f'<!--FORMULA:{len(self.formulas)}-->'

    def _para_text(self, para) -> str:
        """段落文本：含 [^N] 脚注标记与 <!--FORMULA:N--> 公式标记（按原位置）"""
        parts: List[str] = []

        def walk(el):
            for child in el:
                tag = child.tag
                if tag in (q('w:pPr'), q('w:rPr')):
                    continue
                if tag == q('w:r'):
                    parts.append(self._run_text(child))
                elif tag in (q('m:oMath'), q('m:oMathPara')):
                    parts.append(self._register_formula(child))
                else:
                    # hyperlink / ins / sdt / smartTag 等容器：继续下钻
                    walk(child)

        walk(para)
        return ''.join(parts)

    def _para_image_rids(self, para) -> List[str]:
        """段落内图片 rId（按出现顺序去重；含 DrawingML 与 VML 两种）"""
        rids: List[str] = []
        seen = set()
        r_embed = f'{{{NSMAP["r"]}}}embed'
        r_id = f'{{{NSMAP["r"]}}}id'
        for blip in para.findall('.//' + q('a:blip')):
            rid = blip.get(r_embed)
            if rid and rid in self.image_map and rid not in seen:
                seen.add(rid)
                rids.append(rid)
        for imagedata in para.findall('.//' + q('v:imagedata')):
            rid = imagedata.get(r_id)
            if rid and rid in self.image_map and rid not in seen:
                seen.add(rid)
                rids.append(rid)
        return rids

    # ── 结构识别 ──

    def _heading_level(self, para) -> int:
        """标题层级：w:outlineLvl → Heading 样式 → 数字样式 ID；0 = 非标题"""
        pPr = para.find(q('w:pPr'))
        if pPr is None:
            return 0
        outline = pPr.find(q('w:outlineLvl'))
        if outline is not None:
            try:
                lv = int(outline.get(q('w:val'), '9'))
            except ValueError:
                lv = 9
            if lv < 9:
                return min(lv + 1, 3)
        pstyle = pPr.find(q('w:pStyle'))
        style_val = pstyle.get(q('w:val'), '') if pstyle is not None else ''
        if style_val.lower().startswith('heading'):
            return min(int(style_val[-1]), 3) if style_val[-1].isdigit() else 1
        if style_val in ('1', '2', '3'):
            return int(style_val)
        return 0

    def _is_toc_para(self, para) -> bool:
        """目录条目段落：TOC 样式或含 TOC 域"""
        pPr = para.find(q('w:pPr'))
        if pPr is not None:
            pstyle = pPr.find(q('w:pStyle'))
            if pstyle is not None and pstyle.get(q('w:val'), '').lower().startswith('toc'):
                return True
        for instr in para.findall('.//' + q('w:instrText')):
            if instr.text and 'TOC' in instr.text:
                return True
        return False

    def read_cover_fields(self) -> Dict[str, str]:
        """从封面/扉页表格提取元数据字段（同字段首个命中为准），
        另扫段落取「专业代码」。只扫摘要/目录之前的表格，避免正文数据表污染。"""
        fields: Dict[str, str] = {}
        cutoff = len(self.body)
        for i, el in enumerate(self.body):
            if el.tag == q('w:p'):
                norm = _norm(self._plain_text(el))
                if norm in ('摘要', '目录') or norm.lower() == 'abstract':
                    cutoff = i
                    break
        cover_tables = [el for el in list(self.body)[:cutoff] if el.tag == q('w:tbl')]
        for tbl in cover_tables:
            for row in tbl.findall('.//' + q('w:tr')):
                cells = row.findall(q('w:tc'))
                if len(cells) < 2:
                    continue
                label = _norm(self._plain_text(cells[0])).strip('：:')
                value = ''
                for cell in cells[1:]:
                    v = self._plain_text(cell).strip().strip('：:').strip()
                    if v:
                        value = v
                        break
                if not label or not value:
                    continue
                for key_label, meta_key in _COVER_LABEL_MAP:
                    if label == key_label or label.startswith(key_label):
                        fields.setdefault(meta_key, value)
                        break
        if 'majorCode' not in fields:
            for el in list(self.body)[:cutoff]:
                if el.tag != q('w:p'):
                    continue
                m = re.search(r'专\s*业\s*代\s*码\D*(\d{6})', self._plain_text(el))
                if m:
                    fields['majorCode'] = m.group(1)
                    break
        return fields

    def read_blocks(self) -> Tuple[List[Block], Dict[str, int]]:
        """顺序遍历 body 顶层元素，产出 Block 列表 + 锚点索引。

        锚点键：abstract_zh / abstract_en / toc / body_start / ref /
        appendix / ack / attachment（附件N，其后内容全部丢弃）。
        锚点值为 Block 列表中的下标。
        """
        blocks: List[Block] = []
        anchors: Dict[str, int] = {}

        for idx, el in enumerate(self.body):
            tag = el.tag
            if tag == q('w:tbl'):
                table = self._read_table(el)
                if table:
                    blocks.append(Block(kind='table', table=table, index=idx))
                continue
            if tag != q('w:p'):
                continue
            if self._is_toc_para(el):
                continue

            try:
                text = self._para_text(el).strip()
            except Exception as exc:  # 单段容错
                self.warnings.append(f"第 {idx} 段解析失败已跳过：{exc}")
                continue

            image_rids = self._para_image_rids(el)
            if image_rids:
                blocks.append(Block(kind='image', image_rids=image_rids,
                                    text=text, index=idx))
                continue
            if not text:
                continue

            norm = _norm(re.sub(r'<!--FORMULA:\d+-->', '', text))
            level = self._heading_level(el)
            # 守卫：真标题不会太长——有些文档给正文段落也套了带大纲级别的
            # 样式，不拦会把整段综述变成目录条目
            if level > 0 and len(norm) > 60:
                level = 0
                msg = "部分段落带标题样式但内容过长，已按正文处理（不进目录）"
                if msg not in self.warnings:
                    self.warnings.append(msg)

            # 内联标签锚点：摘要/Abstract 与内容同段（如「摘要：随着……」）。
            # 拆成 标签块 + 内容块，下游区域组装逻辑无需感知差异。
            inline = None
            if 'abstract_zh' not in anchors:
                m = re.match(r'^[【\[]?\s*摘\s*要\s*[】\]]?\s*[:：]\s*(\S.*)$',
                             text.strip(), re.S)
                if m:
                    inline = ('abstract_zh', '摘要', m.group(1).strip())
            if inline is None and 'abstract_en' not in anchors:
                m = re.match(r'^\s*ABSTRACT\s*[:：]\s*(\S.*)$', text.strip(),
                             re.IGNORECASE | re.S)
                if m:
                    inline = ('abstract_en', 'Abstract', m.group(1).strip())
            if inline is not None:
                key, label, content = inline
                anchors[key] = len(blocks)
                blocks.append(Block(kind='para', text=label, index=idx))
                blocks.append(Block(kind='para', text=content, index=idx))
                continue

            # 锚点识别（精确匹配特殊节名，容忍【】括号与尾冒号变体）
            plain = norm.strip('【】[]').rstrip(':：')
            if plain in ('摘要', '内容摘要') and 'abstract_zh' not in anchors:
                anchors['abstract_zh'] = len(blocks)
            elif plain.lower() == 'abstract' and 'abstract_en' not in anchors:
                anchors['abstract_en'] = len(blocks)
            elif plain == '目录' and 'toc' not in anchors:
                anchors['toc'] = len(blocks)
            elif plain in ('参考文献', '参考书目') and 'ref' not in anchors:
                anchors['ref'] = len(blocks)
            elif plain == '附录' and 'appendix' not in anchors:
                anchors['appendix'] = len(blocks)
            elif plain in ('致谢', '致辞', '谢辞') and 'ack' not in anchors:
                anchors['ack'] = len(blocks)
            elif norm.startswith('附件') and len(norm) <= 30 and 'attachment' not in anchors:
                anchors['attachment'] = len(blocks)

            if level > 0:
                blocks.append(Block(kind='heading', level=level, text=text, index=idx))
            else:
                blocks.append(Block(kind='para', text=text, index=idx))

        self._detect_body_start(blocks, anchors)
        return blocks, anchors

    def _detect_body_start(self, blocks: List[Block], anchors: Dict[str, int]) -> None:
        """body_start 四级回退探测"""
        after = anchors.get('toc', anchors.get('abstract_en',
                            anchors.get('abstract_zh', -1)))
        end = min((v for k, v in anchors.items()
                   if k in ('ref', 'appendix', 'ack', 'attachment')), default=len(blocks))

        # 1) 首个 level-1 标题
        for i in range(after + 1, end):
            b = blocks[i]
            if b.kind == 'heading' and b.level == 1:
                anchors['body_start'] = i
                return
        # 2) 导论/绪论/引言 文本行或「1 xxx」数字标题行
        for i in range(after + 1, end):
            b = blocks[i]
            if b.kind != 'para':
                continue
            norm = _norm(b.text)
            if (re.match(r'^[一1１]?[、.\s]*(导论|绪论|引言)', norm)
                    or re.match(r'^\d+\s+[一-鿿]', b.text.strip())):
                anchors['body_start'] = i
                self.warnings.append("正文起点靠文本关键词定位（文档标题未使用大纲级别/标题样式）")
                return
        # 3) 目录/摘要之后第一个非空块
        if after >= 0 and after + 1 < end:
            anchors['body_start'] = after + 1
            self.warnings.append("未识别到正文起始标题，正文从摘要/目录后的第一段开始")
            return
        # 4) 文档起点
        anchors['body_start'] = 0
        self.warnings.append("未识别到摘要/目录/正文标题，整篇按正文处理")

    def _read_table(self, tbl) -> Optional[List[List[str]]]:
        """docx 表格 → 二维文本；空表返回 None"""
        try:
            rows = []
            for tr in tbl.findall('.//' + q('w:tr')):
                cells = []
                for tc in tr.findall(q('w:tc')):
                    parts = [self._plain_text(p).strip() for p in tc.findall(q('w:p'))]
                    cells.append(' '.join(t for t in parts if t).replace('|', '\\|'))
                if cells:
                    rows.append(cells)
            if not rows or not any(any(c for c in row) for row in rows):
                return None
            if tbl.findall('.//' + q('a:blip')):
                self.warnings.append("表格内含图片，已忽略（仅保留文字内容）")
            return rows
        except Exception as exc:
            self.warnings.append(f"一处表格解析失败已跳过：{exc}")
            return None

    def export_images(self, blocks: List[Block], photo_dir: Path) -> Dict[str, str]:
        """按正文出现顺序把图片导出到 photo/，命名 1.ext、2.ext…

        Returns:
            rId → 相对文件名（如 "photo/1.png"）；不支持的格式不在其中
        """
        mapping: Dict[str, str] = {}
        counter = 0
        for block in blocks:
            for rid in block.image_rids:
                if rid in mapping:
                    continue
                src = self.image_map.get(rid)
                if not src:
                    continue
                ext = Path(src).suffix.lower()
                if ext not in _IMAGE_EXTS:
                    self.warnings.append(
                        f"跳过不支持的图片格式 {ext}（常见于 Word 剪贴画/矢量图），请转为 png/jpg 后重试")
                    continue
                try:
                    data = self.zf.read(src)
                except KeyError:
                    continue
                counter += 1
                photo_dir.mkdir(parents=True, exist_ok=True)
                name = f"{counter}{ext}"
                (photo_dir / name).write_bytes(data)
                mapping[rid] = f"photo/{name}"
        return mapping

    def export_formulas(self, formula_dir: Path) -> int:
        """把收集到的 OMML 写为 formula/N.xml，返回数量"""
        if not self.formulas:
            return 0
        formula_dir.mkdir(parents=True, exist_ok=True)
        for i, xml in enumerate(self.formulas, 1):
            (formula_dir / f"{i}.xml").write_bytes(xml)
        return len(self.formulas)

    def close(self) -> None:
        self.zf.close()

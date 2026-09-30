"""Markdown 解析器：解析元数据（封面字段）和正文各章节。

从 C# MarkdownParser 直接移植，保持相同的解析逻辑。
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Optional


# ═══════════════════════════════════════════
# 数据结构
# ═══════════════════════════════════════════

# 已知的特殊章节名称（含全角空格变体）
KNOWN_NAMES = [
    "摘 要", "摘要", "Abstract",
    "目 录", "目录",
    "导 论", "导论",
    "参考文献",
    "致 谢", "致谢",
    "附 录", "附录",
]

# 章节类型枚举
SECTION_TYPES = {
    "abstract": "abstract",
    "eng_abstract": "eng_abstract",
    "toc": "toc",
    "references": "references",
    "acknowledgment": "acknowledgment",
    "appendix": "appendix",
    "body": "body",
}

# 本地章节编辑器用它显式覆盖“裸编号/已知节名自动识别”。标记只存在于
# 中间 Markdown；解析时会被剥掉，绝不会进入最终 Word 正文。
BODY_OVERRIDE_PREFIX = "<!--SXUPAPER:BODY-->"


@dataclass
class ParsedSection:
    """解析后的章节"""
    level: int = 0
    title: str = ""
    content: str = ""
    section_type: str = "body"
    tables: List["TableData"] = field(default_factory=list)


@dataclass
class TableData:
    """解析后的表格数据"""
    caption: str = ""          # 表标题，如 "表1-1 数据统计"
    headers: List[str] = field(default_factory=list)
    rows: List[List[str]] = field(default_factory=list)


@dataclass
class ParsedDocument:
    """解析后的完整文档"""
    metadata: Dict[str, str] = field(default_factory=dict)
    sections: List[ParsedSection] = field(default_factory=list)
    footnotes: Dict[int, str] = field(default_factory=dict)   # [^N]: 定义 → {N: 内容}

    @property
    def title(self) -> str:
        """获取论文标题"""
        if self.metadata.get("title"):
            return self.metadata["title"]
        for sec in self.sections:
            if (sec.level >= 1
                    and sec.section_type not in ("abstract", "eng_abstract", "toc",
                                                  "references", "acknowledgment", "appendix")):
                return sec.title
        return "学年论文"

    def get_meta(self, key: str, default: str = "") -> str:
        return self.metadata.get(key, default)

    def find_section(self, keyword: str) -> Optional[ParsedSection]:
        """按关键词查找章节（忽略空格）"""
        kw = keyword.replace(" ", "")
        for sec in self.sections:
            if sec.title.replace(" ", "") == kw or kw in sec.title.replace(" ", ""):
                return sec
        return None

    def get_body_sections(self) -> List[ParsedSection]:
        """获取正文类型章节"""
        return [s for s in self.sections if s.section_type == "body"]

    def get_abstract_content(self) -> str:
        """提取中文摘要正文（去除关键词部分）"""
        sec = self.find_section("摘要")
        if sec is None:
            sec = next((s for s in self.sections if s.section_type == "abstract"), None)
        if sec is None:
            return ""
        c = sec.content
        # 匹配独立的关键词段落
        kw_match = re.search(r'\r?\n\s*关键词[：:\s]*\r?\n.+?(?:\r?\n|$)', c, re.DOTALL)
        if kw_match:
            return _clean_text(c[:kw_match.start()].strip())
        # 备选：从末尾向前查找
        for kw in ("关键词",):
            idx = c.rfind(kw)
            if idx >= 0:
                after = c[idx + len(kw):].strip()
                if '；' in after or ';' in after:
                    return _clean_text(c[:idx].strip())
        return _clean_text(c)

    def get_abstract_keywords(self) -> str:
        """提取中文关键词"""
        sec = self.find_section("摘要")
        if sec is None:
            sec = next((s for s in self.sections if s.section_type == "abstract"), None)
        if sec is None:
            return ""
        c = sec.content
        # 优先匹配独立的关键词段落
        kw_match = re.search(r'关键词[：:\s]*\r?\n(.+?)(?:\r?\n|$)', c, re.DOTALL)
        if kw_match:
            keywords = kw_match.group(1).strip()
            if '；' in keywords or ';' in keywords:
                return _clean_text(keywords)
        # 备选
        for kw in ("关键词：", "关键词:", "关键词"):
            idx = c.rfind(kw)
            if idx >= 0:
                after = c[idx + len(kw):].strip().lstrip('；; \n\r')
                if '；' in after or ';' in after:
                    return _clean_text(after)
        return ""

    def get_eng_abstract_content(self) -> str:
        """提取英文摘要正文"""
        sec = next((s for s in self.sections if s.section_type == "eng_abstract"), None)
        if sec is None:
            return ""
        c = sec.content
        kw_match = re.search(
            r'\r?\n\s*Key\s*words[：:\s]*\r?\n.+?(?:\r?\n|$)',
            c, re.DOTALL | re.IGNORECASE)
        if kw_match:
            return _clean_text(c[:kw_match.start()].strip())
        for kw in ("Keywords", "Key words"):
            idx = c.lower().rfind(kw.lower())
            if idx >= 0:
                after = c[idx + len(kw):].strip()
                if ';' in after:
                    return _clean_text(c[:idx].strip())
        return _clean_text(c)

    def get_eng_keywords(self) -> str:
        """提取英文关键词"""
        sec = next((s for s in self.sections if s.section_type == "eng_abstract"), None)
        if sec is None:
            return ""
        c = sec.content
        kw_match = re.search(
            r'Key\s*words[：:\s]*\r?\n(.+?)(?:\r?\n|$)',
            c, re.DOTALL | re.IGNORECASE)
        if kw_match:
            keywords = kw_match.group(1).strip()
            if ';' in keywords:
                return _clean_text(keywords)
        for kw in ("Keywords:", "Key words:", "Keywords", "Key words"):
            idx = c.lower().rfind(kw.lower())
            if idx >= 0:
                after = c[idx + len(kw):].strip().lstrip(':')
                if ';' in after:
                    return _clean_text(after)
        return ""

    def get_references(self) -> List[str]:
        """提取参考文献列表"""
        sec = next((s for s in self.sections if s.section_type == "references"), None)
        if sec is None:
            return []
        lines = _clean_text(sec.content).split('\n')
        result = []
        for line in lines:
            line = line.strip()
            if not line:
                continue
            if re.match(r'^\[\d+\]', line):
                result.append(line)
            elif re.match(r'^\d+\.\s', line):
                result.append(re.sub(r'^(\d+)\.\s', r'[\1] ', line))
        return result

    def get_acknowledgment_content(self) -> str:
        """提取致谢内容"""
        sec = next((s for s in self.sections if s.section_type == "acknowledgment"), None)
        return _clean_text(sec.content) if sec else ""

    def has_appendix(self) -> bool:
        return any(s.section_type == "appendix" for s in self.sections)


# ═══════════════════════════════════════════
# 辅助函数
# ═══════════════════════════════════════════

def _is_known(line: str) -> bool:
    """判断是否为已知特殊章节名称"""
    norm = line.replace(" ", "")
    for k in KNOWN_NAMES:
        if line == k or norm == k.replace(" ", ""):
            return True
    return False


def _clean_text(text: str, strip_images: bool = True) -> str:
    """清理文本：移除 HTML 注释、br 标签、Markdown 标记等。

    Args:
        text: 待清理文本
        strip_images: 是否移除图片引用（默认 True，用于摘要/致谢等；
                      正文应传 False 以保留图片）
    """
    text = re.sub(r'<!--.*?-->', '', text, flags=re.DOTALL)
    text = re.sub(r'<br\s*/?>', '', text, flags=re.IGNORECASE)
    if strip_images:
        text = re.sub(r'!\[.*?\]\(.*?\)', '', text)
    text = text.replace('\\[', '[').replace('\\]', ']')
    # 移除 Markdown 粗体/斜体标记
    text = re.sub(r'\*{1,3}(.+?)\*{1,3}', r'\1', text)
    # 移除残留的单独 *
    text = re.sub(r'\*+', '', text)
    return text.strip()


def _detect_level(line: str) -> int:
    """检测标题层级（非 # 开头的标题）"""
    trimmed = line.strip()
    if trimmed.startswith('#'):
        lv = 0
        while lv < len(trimmed) and trimmed[lv] == '#':
            lv += 1
        return lv
    if re.match(r'^\d+\.\d+\.\d+\s', trimmed):
        return 3
    if re.match(r'^\d+\.\d+\s', trimmed):
        return 2
    if re.match(r'^\d+\s+[一-鿿]', trimmed) and not re.match(r'^\[\d+\]', trimmed):
        return 1
    return 0


def _strip_numbering(title: str, level: int) -> str:
    """去除标题中的编号"""
    if level <= 0:
        return title
    return re.sub(r'^\d+(\.\d+)*\.?\s+', '', title.strip())


def _classify(title: str, level: int) -> str:
    """根据标题内容判断章节类型"""
    if level == 0:
        return "body"
    norm = title.replace(" ", "")
    if "摘" in norm and "要" in norm:
        return "abstract"
    if "Abstract" in title or "abstract" in title:
        return "eng_abstract"
    if "目录" in norm:
        return "toc"
    if "参考文献" in norm:
        return "references"
    if "致谢" in norm:
        return "acknowledgment"
    if "附录" in norm:
        return "appendix"
    return "body"


# ═══════════════════════════════════════════
# 解析入口
# ═══════════════════════════════════════════


def _parse_table_lines(lines: List[str]) -> Optional[TableData]:
    """解析一组 Markdown 表格行为 TableData。

    Args:
        lines: 连续的 | 开头的表格行（已去除空行）

    Returns:
        TableData 或 None（如果行数不足）
    """
    if len(lines) < 2:
        return None

    def _split_row(line: str) -> List[str]:
        """拆分一行表格单元格"""
        line = line.strip()
        if line.startswith('|'):
            line = line[1:]
        if line.endswith('|'):
            line = line[:-1]
        return [cell.strip() for cell in line.split('|')]

    # 第一行：表头
    headers = _split_row(lines[0])

    # 第二行：分隔行（|---|---|）——跳过
    sep_idx = 1
    if sep_idx < len(lines) and re.match(r'^[\s|:\-]+$', lines[sep_idx]):
        sep_idx = 2

    # 剩余行：数据行
    rows = []
    for i in range(sep_idx, len(lines)):
        row = _split_row(lines[i])
        rows.append(row)

    return TableData(headers=headers, rows=rows)


def parse_markdown(file_path: str | Path) -> ParsedDocument:
    """解析 Markdown 文件，返回 ParsedDocument。

    Args:
        file_path: Markdown 文件路径

    Returns:
        ParsedDocument 包含 metadata 和 sections
    """
    file_path = Path(file_path)
    # utf-8-sig：兼容带 BOM 的 UTF-8 文件（如 Windows 记事本/PowerShell 保存的 md）
    lines = file_path.read_text(encoding="utf-8-sig").splitlines()

    doc = ParsedDocument()
    start = _parse_metadata_block(lines, doc)
    _parse_sections(lines, start, doc)
    return doc


def _parse_metadata_block(lines: List[str], doc: ParsedDocument) -> int:
    """解析元数据块（前10行左右的封面字段）。

    Returns:
        正文开始的行号
    """
    if not lines or lines[0].startswith("#"):
        return 0

    i = 0
    meta_lines: List[str] = []

    # 跳过开头的指令行（如 "请读取这本Markdown文件..."）
    while i < len(lines):
        t = lines[i].strip()
        if not t or t == "---":
            i += 1
            continue
        # 跳过看起来像指令的行（包含"请读取"、"Markdown文件"等）
        if any(kw in t for kw in ('请读取', 'Markdown文件', '转化为', '配合', '格式.md')):
            i += 1
            continue
        break

    for i in range(i, len(lines)):
        t = lines[i].strip()
        if t.startswith("#") or _is_known(t):
            break
        if not t:
            # 空行：如果后面是标题或已知名称，跳过
            if i + 1 < len(lines) and (lines[i + 1].strip().startswith("#") or _is_known(lines[i + 1].strip())):
                i += 1
                break
            if meta_lines:
                i += 1
                break
        if t and t != "---":
            meta_lines.append(t)

    # 跳过分隔符和空行
    while i < len(lines) and (lines[i].strip() == "---" or not lines[i].strip()):
        i += 1

    # 按位置顺序解析封面字段。
    # 单行也解析：用户可能只补填一个字段（如仅题目）；这些行本就会被
    # 从正文里剥离（return i），不解析只会让它们静默丢失。
    if len(meta_lines) >= 1:
        idx = 0
        # 专业代码（6位数字）
        if idx < len(meta_lines) and re.match(r'^\d{6}$', meta_lines[idx]):
            doc.metadata["majorCode"] = meta_lines[idx]
            idx += 1
        # 中文标题（长度 ≥ 8；姓名为 2-6 字，无冲突空间）
        if idx < len(meta_lines) and len(meta_lines[idx]) >= 8:
            doc.metadata["title"] = meta_lines[idx]
            idx += 1
        # 英文标题（纯英文，长度 > 10）
        if (idx < len(meta_lines) and len(meta_lines[idx]) > 10
                and re.match(r'^[A-Za-z\s,:\'\-\(\)\.]+$', meta_lines[idx])):
            doc.metadata["engTitle"] = meta_lines[idx]
            idx += 1
        # 姓名（2-6字符）
        if idx < len(meta_lines) and 2 <= len(meta_lines[idx]) <= 6:
            doc.metadata["name"] = meta_lines[idx]
            idx += 1
        # 学号（12位数字）
        if idx < len(meta_lines) and re.match(r'^\d{12}$', meta_lines[idx]):
            doc.metadata["studentId"] = meta_lines[idx]
            idx += 1
        # 班级
        if idx < len(meta_lines) and ("班" in meta_lines[idx] or "级" in meta_lines[idx]):
            doc.metadata["class"] = meta_lines[idx]
            idx += 1
        # 专业（不含"学院"）
        if idx < len(meta_lines) and "学院" not in meta_lines[idx]:
            doc.metadata["major"] = meta_lines[idx]
            idx += 1
        # 学院
        if idx < len(meta_lines) and "学院" in meta_lines[idx]:
            doc.metadata["college"] = meta_lines[idx]
            idx += 1
        # 指导教师
        if idx < len(meta_lines) and any(k in meta_lines[idx] for k in ("讲", "教授", "师")):
            doc.metadata["advisor"] = meta_lines[idx]
            idx += 1
        # 日期
        if idx < len(meta_lines) and "年" in meta_lines[idx]:
            doc.metadata["date"] = meta_lines[idx]
            idx += 1

    return i


def _parse_sections(lines: List[str], start: int, doc: ParsedDocument) -> None:
    """解析正文各章节"""
    buf: List[str] = []
    tables_buf: List[TableData] = []   # 当前章节的表格列表
    cur_title = ""
    cur_level = 0
    first_h1 = True

    def flush():
        nonlocal cur_title, cur_level, first_h1
        if buf or cur_title:
            content = "\n".join(buf).strip()
            st = _classify(cur_title, cur_level)
            doc.sections.append(ParsedSection(
                level=cur_level, title=cur_title,
                content=content, section_type=st,
                tables=list(tables_buf),
            ))
            if st == "body" and cur_level == 1:
                first_h1 = False
        buf.clear()
        tables_buf.clear()
        cur_title = ""
        cur_level = 0

    i = start
    while i < len(lines):
        line = lines[i].rstrip()
        if line.strip() == "---":
            i += 1
            continue

        # 用户在章节编辑器里明确选择“改为正文”时，优先级必须高于
        # 已知节名和编号标题的自动识别。剥掉内部标记后再作为普通正文保存。
        forced_body = line.startswith(BODY_OVERRIDE_PREFIX)
        if forced_body:
            line = line[len(BODY_OVERRIDE_PREFIX):]

        # # 开头的标题
        if not forced_body and line.strip().startswith("#"):
            flush()
            trimmed = line.strip()
            lv = 0
            while lv < len(trimmed) and trimmed[lv] == '#':
                lv += 1
            cur_title = trimmed[lv:].strip()
            cur_level = lv
            i += 1
            continue

        # 已知特殊章节名称
        if not forced_body and _is_known(line.strip()) and cur_title != line.strip():
            flush()
            cur_title = line.strip()
            cur_level = 1
            i += 1
            continue

        # 自动检测标题层级
        detected = 0 if forced_body else _detect_level(line)
        if detected > 0:
            flush()
            cur_title = line.strip()
            cur_level = detected
            i += 1
            continue

        # 脚注定义行 [^N]: 内容（收集后从正文剔除）
        fn_def = re.match(r'^\[\^(\d+)\]:\s*(.+)$', line.strip())
        if fn_def:
            doc.footnotes[int(fn_def.group(1))] = fn_def.group(2).strip()
            i += 1
            continue

        # 检测表格：连续的 | 开头行
        if line.strip().startswith("|") and "|" in line.strip()[1:]:
            table_lines: List[str] = []
            while i < len(lines) and lines[i].strip().startswith("|"):
                table_lines.append(lines[i].strip())
                i += 1
            table_data = _parse_table_lines(table_lines)
            if table_data is not None:
                # 检查 buf 末尾是否有表标题（如 "表1-1 xxx"、附录里的 "表A-1 xxx"）
                caption = ""
                if buf and re.match(r'^表\s*[A-Za-z]?[-－]?\d', buf[-1].strip()):
                    caption = buf.pop().strip()
                table_data.caption = caption
                tables_buf.append(table_data)
                # 在 content 中插入占位标记
                buf.append(f"<!--TABLE:{len(tables_buf) - 1}-->")
            continue

        buf.append(line)
        i += 1

    flush()

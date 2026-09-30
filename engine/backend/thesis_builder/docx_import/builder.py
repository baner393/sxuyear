"""md 组装器：把 reader 产出的 Block 列表组装成符合本项目契约的 Markdown。

契约基准：parser.parse_markdown 的实际行为（见 parser.py），
样例参照 input/test2/鸟尊论文.md。
"""

from __future__ import annotations

import re
from pathlib import Path
from typing import Dict, List, Optional

from .reader import Block, DocxReader, _norm

# 英文标题谓词（与 parser._parse_metadata_block 完全一致）
_ENG_TITLE_RE = re.compile(r"^[A-Za-z\s,:'\-\(\)\.]+$")

# 全角/花式字符 → parser 谓词可接受的 ASCII
_ENG_NORMALIZE = str.maketrans({
    '：': ': ', '，': ', ', '（': '(', '）': ')',
    '—': '-', '–': '-', '‒': '-', '−': '-',
    '’': "'", '‘': "'", '“': '', '”': '',
    '　': ' ',
})


def build_metadata_block(fields: Dict[str, str], warnings: List[str]) -> List[str]:
    """按 parser 槽位顺序输出元数据行，逐字段用 parser 同款谓词自检。

    不合格字段整行丢弃 + warning（converter 对缺失字段有默认值兜底）。
    """
    lines: List[str] = []

    def drop(key: str, label: str, value: str) -> None:
        warnings.append(f"封面字段「{label}」未能识别（值：{value[:30]}），已留空，可在生成的 md 中手动补充")

    major_code = _norm(fields.get('majorCode', ''))
    if major_code:
        if re.fullmatch(r'\d{6}', major_code):
            lines.append(major_code)
        else:
            drop('majorCode', '专业代码', major_code)

    title = re.sub(r'\s+', ' ', fields.get('title', '')).strip()
    title_ok = len(title) >= 8    # 与 parser 的题目谓词保持一致
    if title and title_ok:
        lines.append(title)
    elif title:
        drop('title', '中文题目', title)

    eng_title = re.sub(r'\s+', ' ', fields.get('engTitle', '').translate(_ENG_NORMALIZE)).strip()
    if eng_title:
        # 若中文题目缺失，engTitle 会被 parser 误认作 title，只能一并丢弃
        if title_ok and len(eng_title) > 10 and _ENG_TITLE_RE.match(eng_title):
            lines.append(eng_title)
        else:
            drop('engTitle', '英文题目', eng_title)

    name = _norm(fields.get('name', ''))
    if name:
        if 2 <= len(name) <= 6:
            lines.append(name)
        else:
            drop('name', '姓名', name)

    student_id = _norm(fields.get('studentId', ''))
    if student_id:
        if re.fullmatch(r'\d{12}', student_id):
            lines.append(student_id)
        else:
            drop('studentId', '学号', student_id)

    class_name = _norm(fields.get('class', ''))
    if class_name:
        if '班' in class_name or '级' in class_name:
            lines.append(class_name)
        else:
            drop('class', '班级', class_name)

    major = _norm(fields.get('major', ''))
    if major:
        if '学院' not in major:
            lines.append(major)
        else:
            drop('major', '专业', major)

    college = _norm(fields.get('college', ''))
    if college:
        if '学院' in college:
            lines.append(college)
        else:
            drop('college', '学院', college)

    # 指导教师：封面常见字符间空格（如「刘春荣 教 授」），先去全部空白
    # 再按职称词重组为「姓名 职称」
    advisor = _norm(fields.get('advisor', ''))
    if advisor:
        m = re.match(r'^(.*?)(副教授|教授|讲师|助教|导师|老师)$', advisor)
        if m and m.group(1):
            advisor = f'{m.group(1)} {m.group(2)}'
        if any(k in advisor for k in ('讲', '教授', '师')):
            lines.append(advisor)
        else:
            drop('advisor', '指导教师', advisor)

    date = _norm(fields.get('date', ''))
    if date:
        if '年' in date:
            lines.append(date)
        else:
            drop('date', '完成时间', date)

    if not lines:
        warnings.append("未从封面识别到任何元数据（姓名/学号/题目等），封面将使用默认占位，可在 md 开头补充")
    return lines


def _normalize_keywords_zh(text: str) -> str:
    after = re.sub(r'^关\s*键\s*词[：:\s]*', '', text.strip())
    after = after.replace('，', '；').replace(',', '；').replace(';', '；')
    after = after.rstrip('。．.')
    return '关键词：' + after


def _normalize_keywords_en(text: str) -> str:
    after = re.sub(r'^key\s*words?[：:\s]*', '', text.strip(), flags=re.IGNORECASE)
    after = after.replace('；', ';').replace('，', ';')
    if ';' not in after:
        after = after.replace(',', ';')
    after = after.rstrip('。．.')
    return 'Key words: ' + after


def _strip_formula_markers(text: str) -> str:
    return re.sub(r'<!--FORMULA:\d+-->', '', text)


def _table_to_md(table: List[List[str]]) -> List[str]:
    """二维文本 → md 管道表（列数按表头补齐）"""
    header = table[0]
    lines = ['| ' + ' | '.join(header) + ' |',
             '| ' + ' | '.join(['---'] * len(header)) + ' |']
    for row in table[1:]:
        row = list(row) + [''] * (len(header) - len(row))
        lines.append('| ' + ' | '.join(row[:len(header)]) + ' |')
    return lines


# 封面残段特征：无摘要/目录锚点的文档里，正文前的段落若长这样就不守恒保留
_COVER_NOISE_RE = re.compile(
    r'^(学校代码|专业代码|学\s*号|姓\s*名|班\s*级|学\s*院|专\s*业|指导教师|职\s*称|'
    r'完成时间|题\s*目|山西财经大学)|学年论文$')


def _block_has_content(b: Block) -> bool:
    if b.kind == 'table':
        return bool(b.table)
    return bool(_norm(_strip_formula_markers(b.text or '')))


def _squash(text: str) -> str:
    """完整性比对用的归一化：去掉所有非字符（空白/标点/管道等格式符）"""
    return re.sub(r'[\W_]+', '', text, flags=re.UNICODE)


def build_markdown(reader: DocxReader, blocks: List[Block], anchors: Dict[str, int],
                   fields: Dict[str, str], image_names: Dict[str, str],
                   warnings: List[str]) -> str:
    """分区组装 md：元数据块 → # 摘要 → # Abstract → # 目录 → 正文 →
    # 参考文献 → # 附录 → # 致谢 → 脚注定义。

    守恒规则：有文字的块要么进 md，要么在刻意丢弃白名单里（封面区/目录条目/
    附件尾件/封底），绝不静默丢失；未归类的按正文保留 + 警告。
    结尾做完整性自检兜底（防未来改动引入丢内容）。
    """
    chunks: List[str] = []

    meta_lines = build_metadata_block(fields, warnings)
    if meta_lines:
        chunks.append('\n'.join(meta_lines))

    total = len(blocks)
    # attachment（附件N）之后的内容全部丢弃（成绩评定表/校徽等模板尾件）
    hard_end = anchors.get('attachment', total)

    claimed: set = set()       # 已被某区域消费（或判定为空）的块
    transformed: set = set()   # 内容被改写合并的块（关键词/图题），自检跳过
    deliberate: set = set()    # 刻意丢弃白名单

    # 白名单①：附件尾件
    deliberate.update(range(hard_end, total))
    # 白名单②：封面/说明/承诺区 = 首个内容锚点（摘要/Abstract/目录）之前的一切
    #（封面数据已进元数据；说明/承诺文字由模板重新提供）。
    # 若这些锚点全缺，则不划封面区——正文前的内容走守恒保留，防止误丢摘要。
    first_content = min((anchors[k] for k in ('abstract_zh', 'abstract_en', 'toc')
                         if k in anchors), default=None)
    if first_content is not None:
        deliberate.update(range(first_content))
    # 白名单③：目录区（目录锚点 → 正文起点）全部是目录条目——无样式的
    # 纯段落条目 reader 认不出，但 converter 会重建目录，照单丢弃
    if 'toc' in anchors:
        deliberate.update(range(anchors['toc'] + 1,
                                max(anchors['toc'] + 1, anchors.get('body_start', 0))))
    # 白名单④：锚点标签块本身（以 # 标题形式重新输出）与封底/封面字样段
    claimed.update(v for v in anchors.values())
    for i, b in enumerate(blocks):
        if b.kind == 'para' and _norm(b.text) in ('封底', '封面'):
            deliberate.add(i)

    def region(start_key: str) -> range:
        """[锚点+1, 下一个锚点) 的块下标区间"""
        if start_key not in anchors:
            return range(0)
        start = anchors[start_key]
        end = min((v for k, v in anchors.items() if v > start), default=total)
        return range(start + 1, min(end, hard_end))

    def emit_abstract(key: str, title: str, kw_probe, kw_normalize, kw_warn: str) -> None:
        paras: List[str] = []
        kw_line: Optional[str] = None
        for i in region(key):
            b = blocks[i]
            if b.kind not in ('para', 'heading'):
                continue    # 摘要区的表格/图片交给守恒兜底
            claimed.add(i)
            text = _strip_formula_markers(b.text).strip()
            if not text:
                continue
            if kw_probe(text):
                kw_line = kw_normalize(text)
                transformed.add(i)   # 分隔符被归一化，自检跳过
                break
            paras.append(text)
        chunks.append(title)
        chunks.extend(paras)
        if kw_line:
            chunks.append(kw_line)
        else:
            warnings.append(kw_warn)

    # ── 中文摘要 ──
    if 'abstract_zh' in anchors:
        emit_abstract('abstract_zh', '# 摘要',
                      lambda t: '关键词' in _norm(t)[:6], _normalize_keywords_zh,
                      "未识别到中文关键词行（应为「关键词：a；b；c」）")
    else:
        warnings.append("未识别到中文摘要（缺少「摘要」标题段落）")

    # ── 英文摘要 ──
    if 'abstract_en' in anchors:
        emit_abstract('abstract_en', '# Abstract',
                      lambda t: bool(re.match(r'^key\s*words?', t, re.IGNORECASE)),
                      _normalize_keywords_en,
                      "未识别到英文关键词行（应为「Key words: x; y; z」）")
    else:
        warnings.append("未识别到英文摘要（缺少「Abstract」标题段落）")

    # 目录恒输出（converter 自建目录域，条目已在提取时丢弃）
    chunks.append('# 目录')

    # ── 正文 ──
    body_start = anchors.get('body_start', 0)
    body_end = min((v for k, v in anchors.items()
                    if k in ('ref', 'appendix', 'ack') and v >= body_start),
                   default=total)
    body_end = min(body_end, hard_end)
    # 先单独输出正文起始块（标题），守恒保留的前置内容紧随其后——
    # 不能放在「# 目录」之后（parser 会把目录节内容全部丢弃）
    _emit_content(blocks, range(body_start, min(body_start + 1, body_end)), chunks,
                  image_names, warnings, as_body=True, force_heading_at=body_start,
                  transformed=transformed)
    claimed.update(range(body_start, body_end))

    # 守恒兜底（正文前）：没被摘要区认领、又不在白名单的前置内容
    front_kept = _emit_unclaimed(blocks, range(0, body_start), claimed, deliberate,
                                 fields, chunks, image_names, warnings, transformed,
                                 cover_filter=(first_content is None))
    _emit_content(blocks, range(min(body_start + 1, body_end), body_end), chunks,
                  image_names, warnings, as_body=True, transformed=transformed)

    # ── 参考文献 ──
    if 'ref' in anchors:
        entries: List[str] = []
        for i in region('ref'):
            b = blocks[i]
            if b.kind not in ('para', 'heading'):
                break   # 遇图片/表格即止（模板尾部的校徽/评定表）
            claimed.add(i)
            transformed.add(i)   # 条目会被统一重编号，完整性自检跳过
            text = _strip_formula_markers(b.text).strip()
            if not text or _norm(text) in ('封底', '封面'):
                continue
            m = re.match(r'^\[(\d+)\]\s*(.*)', text)
            m2 = re.match(r'^(\d+)[.、]\s*(.*)', text) if not m else None
            if m:
                entries.append(m.group(2).strip())
            elif m2:
                entries.append(m2.group(2).strip())
            elif entries:
                entries[-1] += text          # 无编号行视为上一条的续行
            else:
                entries.append(text)
                warnings.append("参考文献首条缺少 [1] 编号，已自动补编")
        if entries:
            chunks.append('# 参考文献')
            chunks.append('\n'.join(f'[{i}] {e}' for i, e in enumerate(entries, 1)))
        else:
            warnings.append("参考文献一节为空")
    else:
        warnings.append("未识别到参考文献")

    # ── 附录 ──
    if 'appendix' in anchors:
        chunks.append('# 附录')
        _emit_content(blocks, region('appendix'), chunks, image_names,
                      warnings, as_body=False, transformed=transformed)
        claimed.update(region('appendix'))

    # ── 致谢 ──
    if 'ack' in anchors:
        paras = []
        for i in region('ack'):
            b = blocks[i]
            if b.kind == 'para':
                claimed.add(i)
                text = _strip_formula_markers(b.text).strip()
                if text:
                    paras.append(text)
            else:
                break   # 遇表格/图片即止（模板尾部的评定表/校徽）
        chunks.append('# 致谢')
        chunks.extend(paras)

    # 守恒兜底（正文后）：各区域断掉后剩下的内容，补在正文区末尾
    #（此时 chunks 已含后续节标题，插到参考文献标题之前）
    tail_indices = [i for i in range(body_end, total)]
    tail_chunks: List[str] = []
    tail_kept = _emit_unclaimed(blocks, tail_indices, claimed, deliberate,
                                fields, tail_chunks, image_names, warnings,
                                transformed, cover_filter=False)
    if tail_chunks:
        insert_at = next((k for k, c in enumerate(chunks)
                          if c in ('# 参考文献', '# 附录', '# 致谢')), len(chunks))
        chunks[insert_at:insert_at] = tail_chunks

    kept = front_kept + tail_kept
    if kept:
        previews = '；'.join(t[:20] for t in kept[:3])
        warnings.append(
            f"有 {len(kept)} 处内容未能自动归类，已按正文保留（可在导入检查与修正的"
            f"「导入检查与修正」中调整位置或删除）：{previews}…")

    # ── 脚注定义 ──
    if reader.footnotes:
        chunks.append('\n'.join(f'[^{n}]: {t}' for n, t in sorted(reader.footnotes.items())))

    if reader.formulas:
        warnings.append(
            f"文档含 {len(reader.formulas)} 处公式，已按原样保留（md 中为 <!--FORMULA:N--> 占位，"
            "只能整体移动或删除；修改公式内容请回 Word 编辑后重新导入）")

    md_text = '\n\n'.join(chunks) + '\n'

    # ── 完整性自检（保险丝）：白名单之外的每个内容块都必须出现在 md 里 ──
    # 两侧都剥掉公式占位符再比对（块探针剥了，md 不剥会在行内公式处错位）
    md_squash = _squash(_strip_formula_markers(md_text))
    missing: List[str] = []
    for i, b in enumerate(blocks):
        if i in deliberate or i in transformed or not _block_has_content(b):
            continue
        if b.kind == 'table':
            probe = _squash(''.join(b.table[0]))[:20]
        else:
            probe = _squash(_strip_formula_markers(b.text))[:30]
        if len(probe) >= 4 and probe not in md_squash:
            missing.append((b.text or ''.join(b.table[0]))[:24])
    if missing:
        warnings.append(
            f"完整性自检：有 {len(missing)} 处内容未进入提取结果，请检查并反馈："
            + '；'.join(missing[:3]))

    return md_text


def _emit_unclaimed(blocks: List[Block], indices, claimed: set, deliberate: set,
                    fields: Dict[str, str], chunks: List[str],
                    image_names: Dict[str, str], warnings: List[str],
                    transformed: set, cover_filter: bool) -> List[str]:
    """守恒兜底：把未被任何区域认领的内容块按正文形式输出。

    cover_filter=True（文档无摘要/目录锚点）时过滤疑似封面残段：
    命中封面标签特征、或与已识别的元数据值重合的段落仍旧丢弃。
    Returns: 保留内容的文本预览列表（用于警告）。
    """
    field_values = {_norm(v) for v in fields.values() if v}
    kept: List[str] = []
    emit_ids: List[int] = []
    for i in indices:
        if i in claimed or i in deliberate:
            continue
        b = blocks[i]
        if not _block_has_content(b):
            claimed.add(i)
            continue
        if b.kind == 'image' and not any(r in image_names for r in b.image_rids):
            deliberate.add(i)   # 未导出的装饰图（封面校徽等）
            continue
        if cover_filter and b.kind in ('para', 'heading'):
            norm = _norm(_strip_formula_markers(b.text))
            if (_COVER_NOISE_RE.search(norm) and len(norm) <= 40) or norm in field_values:
                deliberate.add(i)
                continue
        emit_ids.append(i)
    if not emit_ids:
        return kept
    # 标题一律降为正文输出：未归类的"标题"多为误判，直接进目录会污染结构，
    # 用户可在“导入检查与修正”的章节结构编辑器里重新标级
    temp = [Block(kind='para', text=blocks[i].text, index=blocks[i].index)
            if blocks[i].kind == 'heading' else blocks[i] for i in emit_ids]
    local_transformed: set = set()
    _emit_content(temp, range(len(temp)), chunks, image_names, warnings,
                  as_body=False, transformed=local_transformed)
    for local_idx in local_transformed:
        transformed.add(emit_ids[local_idx])
    for i in emit_ids:
        claimed.add(i)
        b = blocks[i]
        if b.kind == 'table':
            kept.append('[表格] ' + ''.join(b.table[0])[:20])
        elif _strip_formula_markers(b.text or '').strip():
            kept.append(_strip_formula_markers(b.text).strip())
        else:
            kept.append('[图片]')
    return kept


def _emit_content(blocks: List[Block], indices, chunks: List[str],
                  image_names: Dict[str, str], warnings: List[str],
                  as_body: bool, force_heading_at: Optional[int] = None,
                  transformed: Optional[set] = None) -> None:
    """通用内容区输出：标题/段落/图片(合并图题)/表格(合并表题)/公式

    force_heading_at: 该下标的段落强制按一级标题输出（正文起点靠文本
    关键词定位时它没有标题样式，不升级会落进目录节被丢弃）。
    transformed: 若给出，内容被改写合并的块下标（图题并入图片）记录于此，
    供完整性自检跳过。
    """
    idx_list = list(indices)
    skip: set = set()
    for pos, i in enumerate(idx_list):
        if i in skip:
            continue
        b = blocks[i]

        if b.kind == 'para' and i == force_heading_at:
            chunks.append('# ' + _strip_formula_markers(b.text).strip())
            continue

        if b.kind == 'heading':
            level = min(max(b.level, 1), 3)
            chunks.append('#' * level + ' ' + _strip_formula_markers(b.text).strip())
            continue

        if b.kind == 'image':
            caption = ''
            # 图题在图下方：下一块是「图X-X 描述」段落则并入并消费
            if pos + 1 < len(idx_list):
                nxt = blocks[idx_list[pos + 1]]
                if nxt.kind == 'para' and re.match(r'^图\s*[A-Za-z]?[-－]?[\d\-－—]+', nxt.text.strip()):
                    caption = re.sub(r'^图\s*[A-Za-z]?[-－]?[\d\-－—.]+\s*', '', nxt.text.strip()).strip()
                    skip.add(idx_list[pos + 1])
                    if transformed is not None:
                        transformed.add(idx_list[pos + 1])
            for rid in b.image_rids:
                name = image_names.get(rid)
                if name:
                    chunks.append(f'![{caption}]({name})')
                    caption = ''   # 多图共段时图题只挂第一张
            # 图片段落若还带文字（罕见），文字保留为普通段
            residual = _strip_formula_markers(b.text).strip()
            if residual:
                chunks.append(residual)
            continue

        if b.kind == 'table':
            table_lines = _table_to_md(b.table)
            # 表题在表上方：上一 chunk 是「表X-X」行则合并进同一 chunk
            # （parser 要求表题行与管道表之间无空行；附录里可能是「表A-1」）
            if chunks and re.match(r'^表\s*[A-Za-z]?[-－]?\d', chunks[-1].split('\n')[-1]):
                caption_chunk = chunks.pop()
                chunks.append(caption_chunk + '\n' + '\n'.join(table_lines))
            else:
                chunks.append('\n'.join(table_lines))
            continue

        # para
        text = b.text.strip()
        if not text:
            continue
        # 纯公式段：每个公式标记独立成行（converter 按独立公式居中排版）
        stripped = _strip_formula_markers(text).strip()
        markers = re.findall(r'<!--FORMULA:\d+-->', text)
        if markers and not stripped:
            for mk in markers:
                chunks.append(mk)
            continue
        # 以句读符号开头的段落是上一段被硬拆的残片（脚注后回车等），并回上一段
        if (chunks and text[0] in '。．，、；：？！）'
                and not chunks[-1].startswith(('#', '|', '!['))
                and '\n' not in chunks[-1]):
            chunks[-1] += text
            continue
        chunks.append(text)

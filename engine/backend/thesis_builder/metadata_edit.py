"""封面元数据的表单化写回：把用户填写的字段安全地写进 md 元数据块。

parser._parse_metadata_block 是**顺序槽位匹配**：每行按 专业代码→题目→英文题目→
姓名→学号→班级→专业→学院→导师→日期 的顺序用谓词认领。因此写回必须：
1. 逐字段用与 parser 完全一致的谓词校验，不合格的字段不写（会污染后续槽位），
   原因通过 rejected 返回给前端展示；
2. 按规范顺序输出，允许任意字段缺席（对应槽位自然跳过）；
3. 英文题目依赖中文题目在场（否则题目槽位会误吞英文行）。
"""

from __future__ import annotations

import re
from typing import Dict, List, Tuple

# 规范顺序（= parser 槽位顺序）
META_ORDER = ['majorCode', 'title', 'engTitle', 'name', 'studentId',
              'class', 'major', 'college', 'advisor', 'date']

_ADVISOR_TITLE_RE = re.compile(r'^(.*?)(副教授|教授|讲师|助教|导师|老师)$')
# 英文题目里常见的全角符号 → 半角（与 docx_import.builder 的归一化一致）
_ENG_NORMALIZE = {'：': ':', '，': ',', '’': "'", '‘': "'", '（': '(', '）': ')',
                  '—': '-', '－': '-', '．': '.'}


def _validate(key: str, value: str) -> Tuple[str, str]:
    """校验并归一化单个字段。返回 (归一化后的值, 错误信息)；错误信息为空 = 合格。"""
    v = ' '.join(value.split())  # 压缩内部空白、去首尾
    if key == 'majorCode':
        v = v.replace(' ', '')
        if not re.fullmatch(r'\d{6}', v):
            return v, '专业代码需为 6 位数字'
    elif key == 'title':
        if len(v) < 8:
            return v, '中文题目至少 8 个字（识别约束，过短会与姓名混淆）'
    elif key == 'engTitle':
        for full, half in _ENG_NORMALIZE.items():
            v = v.replace(full, half)
        if len(v) <= 10 or not re.fullmatch(r"[A-Za-z\s,:'\-\(\)\.]+", v):
            return v, '英文题目需为英文字母与常用标点（不含数字），且长于 10 字符'
    elif key == 'name':
        v = v.replace(' ', '')
        if not (2 <= len(v) <= 6):
            return v, '姓名需 2-6 个字符'
    elif key == 'studentId':
        v = v.replace(' ', '')
        if not re.fullmatch(r'\d{12}', v):
            return v, '学号需为 12 位数字'
    elif key == 'class':
        if '班' not in v and '级' not in v:
            return v, '班级需含「班」或「级」，如：文化产业管理2301班'
    elif key == 'major':
        if '学院' in v:
            return v, '专业里不能含「学院」（学院请填在学院栏）'
    elif key == 'college':
        if '学院' not in v:
            return v, '学院需含「学院」二字'
    elif key == 'advisor':
        compact = v.replace(' ', '')
        m = _ADVISOR_TITLE_RE.match(compact)
        if m and m.group(1):
            v = f'{m.group(1)} {m.group(2)}'   # 统一为「姓名 职称」
        else:
            return v, '请填「姓名 职称」，职称为 教授/副教授/讲师 等，如：王某某 讲师'
    elif key == 'date':
        if '年' not in v:
            return v, '完成时间需含「年」，如：2026年5月'
    return v, ''


def build_block(fields: Dict[str, str]) -> Tuple[List[str], Dict[str, str], Dict[str, str]]:
    """把用户填写的字段整理成元数据块行。

    Returns:
        (块内各行, 实际采用的字段, 被拒字段→原因)
    """
    applied: Dict[str, str] = {}
    rejected: Dict[str, str] = {}
    for key in META_ORDER:
        raw = (fields.get(key) or '').strip()
        if not raw:
            continue  # 留空 = 跳过，完全允许
        v, err = _validate(key, raw)
        if err:
            rejected[key] = err
        else:
            applied[key] = v
    # 依赖规则：英文题目无中文题目护航会被 parser 误认成中文题目
    if 'engTitle' in applied and 'title' not in applied:
        rejected['engTitle'] = '需先填写合格的中文题目，英文题目才能被正确识别'
        del applied['engTitle']
    lines = [applied[k] for k in META_ORDER if k in applied]
    return lines, applied, rejected


def _locate_block(lines: List[str]) -> Tuple[int, int]:
    """定位现有元数据块 [start, end)（与 parser 的跳过/终止逻辑一致）。
    无块（直接以 # 开头）时返回 (insert_pos, insert_pos)。"""
    i = 0
    while i < len(lines):
        t = lines[i].strip()
        if not t or t == '---' or any(kw in t for kw in
                ('请读取', 'Markdown文件', '转化为', '配合', '格式.md')):
            i += 1
            continue
        break
    start = i
    if start >= len(lines) or lines[start].strip().startswith('#'):
        return start, start
    end = start
    while end < len(lines):
        t = lines[end].strip()
        if not t or t.startswith('#'):
            break
        end += 1
    return start, end


def update_metadata_block(md_text: str, fields: Dict[str, str]
                          ) -> Tuple[str, Dict[str, str], Dict[str, str]]:
    """用用户填写的字段替换 md 的元数据块。

    Returns:
        (新 md 文本, 采用的字段, 被拒字段→原因)
    """
    block, applied, rejected = build_block(fields)
    lines = md_text.splitlines()
    start, end = _locate_block(lines)
    new_lines = lines[:start] + block
    rest = lines[end:]
    # 块与正文之间保证恰好一个空行
    if block and (not rest or rest[0].strip()):
        new_lines.append('')
    new_lines += rest
    return '\n'.join(new_lines) + '\n', applied, rejected

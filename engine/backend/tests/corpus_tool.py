"""回归语料库收录工具：脱敏 + 生成 expected.md 草稿。

用法（项目根目录）:
    收录新问题文档（自动脱敏姓名/学号/导师姓名/文档作者属性）:
        .venv\\Scripts\\python -m backend.tests.corpus_tool add <某文档.docx> --note "哪里识别错了"

    识别规则有意变更后，重新生成某个/全部 case 的期望结果（会打印 diff，需人工过目）:
        .venv\\Scripts\\python -m backend.tests.corpus_tool refresh [case-001-xxx]

维护流程详见 backend/tests/corpus/README.md。
"""

from __future__ import annotations

import argparse
import re
import shutil
import sys
import tempfile
import zipfile
from pathlib import Path
from typing import Dict, List, Tuple

from lxml import etree

_ROOT = Path(__file__).resolve().parents[2]
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

from backend.thesis_builder.docx_import import extract_docx_to_md  # noqa: E402
from backend.thesis_builder.docx_import.reader import DocxReader, q, _norm  # noqa: E402

CORPUS_DIR = Path(__file__).parent / 'corpus'
_XML_SPACE = '{http://www.w3.org/XML/1998/namespace}space'
# 含 w:t 文本、可能出现人名/学号的部件
_TEXT_PART_RE = re.compile(
    r'^word/(document\.xml|header\d*\.xml|footer\d*\.xml|footnotes\.xml|endnotes\.xml)$'
)
_ADVISOR_TITLE_RE = re.compile(r'^(.*?)(副教授|教授|讲师|助教|导师|老师)$')
# 1x1 透明 PNG：介质瘦身占位符。提取管线只关心图片的 rId 顺序与扩展名，
# 不读像素内容，因此语料 case 里可用占位字节替换真实图片给仓库减重。
_TINY_PNG = bytes.fromhex(
    '89504e470d0a1a0a0000000d49484452000000010000000108060000001f15c489'
    '0000000d4944415478da63f8ffff3f0300050201cfa06b6e0000000049454e44ae426082'
)
_SHRINK_THRESHOLD = 3 * 1024 * 1024  # 超过 3MB 的 docx 默认瘦身介质


# ── 脱敏 ──

def _build_scrub_map(fields: Dict[str, str]) -> Dict[str, str]:
    """从封面字段生成 真实值→同长度假值 的替换表（保留长度等格式特征）"""
    scrub: Dict[str, str] = {}
    name = _norm(fields.get('name', ''))
    if name:
        scrub[name] = '张' + '某' * (len(name) - 1)
    sid = _norm(fields.get('studentId', ''))
    if re.fullmatch(r'\d{8,}', sid):
        scrub[sid] = sid[:4] + '0' * (len(sid) - 4)
    advisor = _norm(fields.get('advisor', ''))
    if advisor:
        m = _ADVISOR_TITLE_RE.match(advisor)
        adv_name = m.group(1) if m and m.group(1) else advisor
        if adv_name and adv_name not in scrub:
            scrub[adv_name] = '王' + '某' * (len(adv_name) - 1)
    return scrub


def _replace_across_runs(para, old: str, new: str) -> int:
    """段内跨 run 替换：w:t 拼接后查找，命中区间回写到各节点。返回替换次数。"""
    wts = para.findall('.//' + q('w:t'))
    if not wts:
        return 0
    texts = [t.text or '' for t in wts]
    combined = ''.join(texts)
    count = 0
    while old in combined:
        start = combined.index(old)
        end = start + len(old)
        pos = 0
        replaced = False
        for i in range(len(texts)):
            n_start, n_end = pos, pos + len(texts[i])
            pos = n_end
            if n_end <= start or n_start >= end:
                continue
            head = texts[i][:max(0, start - n_start)]
            tail = texts[i][max(0, min(len(texts[i]), end - n_start)):]
            texts[i] = head + (new if not replaced else '') + tail
            replaced = True
        combined = ''.join(texts)
        count += 1
    if count:
        for t, txt in zip(wts, texts):
            if (t.text or '') != txt:
                t.text = txt
                t.set(_XML_SPACE, 'preserve')
    return count


def _scrub_core_props(data: bytes) -> bytes:
    """清空 docProps/core.xml 的作者/最后修改者（常带真实姓名或账号）"""
    root = etree.fromstring(data)
    for tag in ('{http://purl.org/dc/elements/1.1/}creator',
                '{http://schemas.openxmlformats.org/package/2006/metadata/core-properties}lastModifiedBy'):
        for el in root.iter(tag):
            el.text = ''
    return etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)


def scrub_docx(src: Path, dst: Path, shrink_media: bool = False) -> Tuple[Dict[str, str], List[str]]:
    """脱敏复制 src→dst。返回 (替换表, 残留告警列表)。

    替换表为空说明封面字段没识别出来，无法自动脱敏（需人工在 Word 里处理）。
    shrink_media=True 时把 word/media/ 下的图片全部换成 1x1 占位 PNG
    （不影响提取结果，只为给语料仓库减重；input.docx 在 Word 里图片会显示为损坏）。
    """
    reader = DocxReader(src)
    try:
        fields = reader.read_cover_fields()
    finally:
        reader.close()
    scrub = _build_scrub_map(fields)
    # 长值优先替换，防止短值先命中拆散长值
    ordered = sorted(scrub.items(), key=lambda kv: -len(kv[0]))

    residuals: List[str] = []
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if shrink_media and item.filename.startswith('word/media/'):
                data = _TINY_PNG
            elif item.filename == 'docProps/core.xml':
                data = _scrub_core_props(data)
            elif _TEXT_PART_RE.match(item.filename) and ordered:
                root = etree.fromstring(data)
                for para in root.iter(q('w:p')):
                    for old, new in ordered:
                        _replace_across_runs(para, old, new)
                # 残留检查：拼全文查真实值（跨段拆分等极端情况）
                full = ''.join(t.text or '' for t in root.iter(q('w:t')))
                for old, _ in ordered:
                    if old in full:
                        residuals.append(f'{item.filename} 中仍残留「{old}」，请在 Word 里手工替换后重新收录')
                data = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
            zout.writestr(item, data)
    return scrub, residuals


# ── 收录 / 刷新 ──

def _next_case_dir(slug: str) -> Path:
    CORPUS_DIR.mkdir(parents=True, exist_ok=True)
    nums = [int(m.group(1)) for d in CORPUS_DIR.iterdir()
            if d.is_dir() and (m := re.match(r'^case-(\d+)', d.name))]
    n = max(nums, default=0) + 1
    slug = re.sub(r'[^\w一-鿿-]+', '-', slug).strip('-') or 'doc'
    return CORPUS_DIR / f'case-{n:03d}-{slug}'


def _extract_expected(case_dir: Path) -> List[str]:
    """对 case 的 input.docx 跑提取，写 expected.md / expected_warnings.txt，返回警告"""
    with tempfile.TemporaryDirectory() as tmp:
        result = extract_docx_to_md(case_dir / 'input.docx', Path(tmp))
        md_text = result.md_path.read_text(encoding='utf-8-sig')
    (case_dir / 'expected.md').write_text(md_text, encoding='utf-8', newline='\n')
    (case_dir / 'expected_warnings.txt').write_text(
        '\n'.join(result.warnings) + ('\n' if result.warnings else ''),
        encoding='utf-8', newline='\n')
    return result.warnings


def add_case(src: Path, name: str = None, note: str = None,
             keep_media: bool = False) -> dict:
    """收录一份问题文档为语料 case（脱敏 + 生成 expected.md 草稿）。

    可被 CLI 与 /api/corpus/add 共用。
    Returns: {case: 目录名, scrubbed: 替换数, residuals: [...], warnings: [...],
              shrunk: 是否瘦身了图片}
    """
    src = Path(src)
    if not src.is_file():
        raise FileNotFoundError(f'找不到文件: {src}')
    case_dir = _next_case_dir(name or src.stem)
    case_dir.mkdir(parents=True)
    shrink = (not keep_media) and src.stat().st_size > _SHRINK_THRESHOLD
    try:
        scrub, residuals = scrub_docx(src, case_dir / 'input.docx', shrink_media=shrink)
        warnings = _extract_expected(case_dir)
        (case_dir / 'note.md').write_text(
            f'# {case_dir.name}\n\n'
            f'**问题**：{note or "（填写：这份文档暴露了什么识别问题）"}\n\n'
            # 注意：note.md 会进 git，绝不能写真实值→假值映射，只记录脱敏了哪些字段
            f'**脱敏**：{"已自动替换 " + str(len(scrub)) + " 个隐私值（姓名/学号/导师姓名，同长度假值）" if scrub else "封面字段未识别，未能自动脱敏——需人工确认无隐私信息"}\n',
            encoding='utf-8', newline='\n')
    except Exception:
        shutil.rmtree(case_dir, ignore_errors=True)
        raise
    return {'case': case_dir.name, 'scrub': scrub, 'residuals': residuals,
            'warnings': warnings, 'shrunk': shrink}


def cmd_add(args: argparse.Namespace) -> int:
    try:
        result = add_case(Path(args.docx), name=args.name, note=args.note,
                          keep_media=args.keep_media)
    except FileNotFoundError as exc:
        print(exc)
        return 1
    if result['shrunk']:
        print('源文件 > 3MB，已自动瘦身图片为占位符（--keep-media 可保留）')
    print(f"已收录: backend/tests/corpus/{result['case']}")
    print(f"脱敏替换: {result['scrub'] or '（无——封面字段未识别，请人工检查 input.docx！）'}")
    for r in result['residuals']:
        print(f'[!] {r}')
    for w in result['warnings']:
        print(f'[提取警告] {w}')
    print('\n下一步（人工）:')
    print(f"  1. 打开 {result['case']}/expected.md，把识别错的地方改成正确答案")
    print('  2. 在 note.md 里写清这份文档暴露的问题')
    print('  3. 跑 pytest backend/tests —— 新 case 应当红（等待规则修复）或绿（纯收录）')
    return 0


def cmd_refresh(args: argparse.Namespace) -> int:
    import difflib
    targets = ([CORPUS_DIR / args.case] if args.case
               else sorted(d for d in CORPUS_DIR.iterdir() if d.is_dir()))
    for case_dir in targets:
        if not (case_dir / 'input.docx').exists():
            print(f'跳过 {case_dir.name}（无 input.docx）')
            continue
        old = ((case_dir / 'expected.md').read_text(encoding='utf-8-sig').splitlines()
               if (case_dir / 'expected.md').exists() else [])
        _extract_expected(case_dir)
        new = (case_dir / 'expected.md').read_text(encoding='utf-8-sig').splitlines()
        diff = list(difflib.unified_diff(old, new, 'expected.md(旧)', 'expected.md(新)', lineterm=''))
        print(f'== {case_dir.name}: {"无变化" if not diff else f"{len(diff)} 行 diff，请人工确认"} ==')
        for line in diff[:80]:
            print(line)
    print('\nrefresh 只是重新生成，正确性仍需人工确认后提交。')
    return 0


def main() -> int:
    ap = argparse.ArgumentParser(description='回归语料库收录/维护工具')
    sub = ap.add_subparsers(dest='cmd', required=True)
    p_add = sub.add_parser('add', help='脱敏收录一份问题文档')
    p_add.add_argument('docx', help='问题 .docx 路径')
    p_add.add_argument('--name', help='case 目录名后缀（默认取文件名）')
    p_add.add_argument('--note', help='一句话：这份文档暴露了什么问题')
    p_add.add_argument('--keep-media', action='store_true',
                       help='保留原始图片字节（默认 >3MB 的文档自动换成占位符减重）')
    p_add.set_defaults(func=cmd_add)
    p_ref = sub.add_parser('refresh', help='规则有意变更后重新生成期望结果（打印 diff）')
    p_ref.add_argument('case', nargs='?', help='case 目录名（缺省=全部）')
    p_ref.set_defaults(func=cmd_refresh)
    args = ap.parse_args()
    return args.func(args)


if __name__ == '__main__':
    sys.exit(main())

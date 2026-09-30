"""回归语料库测试：每个 case 目录 = 一份真实（脱敏）问题文档 + 人工确认的正确提取结果。

运行（项目根目录）:
    .venv\\Scripts\\python -m pytest backend/tests -q

新增 case 用收录命令（见 backend/tests/corpus/README.md）:
    .venv\\Scripts\\python -m backend.tests.corpus_tool add <某文档.docx> --note "问题描述"
"""

import difflib
from pathlib import Path

import pytest

CORPUS_DIR = Path(__file__).parent / 'corpus'


def _collect_cases():
    if not CORPUS_DIR.is_dir():
        return []
    return sorted(
        d for d in CORPUS_DIR.iterdir()
        if d.is_dir() and (d / 'input.docx').exists()
    )


CASES = _collect_cases()


@pytest.mark.parametrize('case', CASES, ids=lambda d: d.name)
def test_corpus_extraction(case: Path, tmp_path: Path) -> None:
    """input.docx 的提取结果必须与人工确认过的 expected.md 逐行一致。"""
    from backend.thesis_builder.docx_import import extract_docx_to_md

    result = extract_docx_to_md(case / 'input.docx', tmp_path)

    actual = result.md_path.read_text(encoding='utf-8-sig').splitlines()
    expected = (case / 'expected.md').read_text(encoding='utf-8-sig').splitlines()
    if actual != expected:
        diff = '\n'.join(difflib.unified_diff(
            expected, actual,
            fromfile=f'{case.name}/expected.md（人工确认）',
            tofile='本次提取结果',
            lineterm='',
        ))
        pytest.fail(f'{case.name} 提取结果与期望不一致：\n{diff}', pytrace=False)

    warn_file = case / 'expected_warnings.txt'
    expected_warnings = (
        warn_file.read_text(encoding='utf-8').splitlines()
        if warn_file.exists() else []
    )
    assert result.warnings == expected_warnings, (
        f'{case.name} 提取警告变化：\n'
        f'期望: {expected_warnings}\n实际: {result.warnings}'
    )


def test_corpus_not_empty() -> None:
    """语料库必须至少有一个 case（防止目录被误删后测试静默全过）。"""
    assert CASES, f'语料库为空：{CORPUS_DIR} 下没有任何 case 目录'

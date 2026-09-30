from __future__ import annotations

from pathlib import Path

import pytest

from backend.thesis_builder.outline_edit import (
    OutlineEditError,
    apply_outline_changes,
    build_outline_snapshot,
)
from backend.thesis_builder.parser import parse_markdown


SAMPLE = """论文标题至少需要八个汉字\n\n# 第一章 绪论\n正文一。\n## 1.1 研究背景\n正文二[1]。\n1.2 遗漏标题\n正文三。\n### 1.1.1 三级标题\n正文四。\n# 参考文献\n[1] 示例文献。\n"""


def write_sample(tmp_path: Path) -> Path:
    path = tmp_path / "paper.md"
    path.write_text(SAMPLE, encoding="utf-8", newline="\n")
    return path


def test_snapshot_separates_recognized_headings_and_candidates(tmp_path):
    path = write_sample(tmp_path)
    snapshot = build_outline_snapshot(path)
    assert [item["text"] for item in snapshot["headings"]] == [
        "第一章 绪论", "1.1 研究背景", "1.2 遗漏标题", "1.1.1 三级标题", "参考文献"
    ]
    assert snapshot["lines"][0]["disabled_reason"] == "metadata"
    assert snapshot["lines"][3]["text"] == "正文一。"
    assert snapshot["source_sha256"]


def test_apply_multiple_levels_is_atomic_and_preserves_nonselected_lines(tmp_path):
    path = write_sample(tmp_path)
    original = path.read_text(encoding="utf-8").splitlines()
    snapshot = build_outline_snapshot(path)
    result = apply_outline_changes(path, snapshot["source_sha256"], [
        {"line_index": 4, "level": 1},
        {"line_index": 6, "level": 2},
    ])
    updated = path.read_text(encoding="utf-8").splitlines()
    assert updated[4] == "# 1.1 研究背景"
    assert updated[6] == "## 1.2 遗漏标题"
    assert all(updated[i] == line for i, line in enumerate(original) if i not in {4, 6})
    assert result["source_sha256"] != snapshot["source_sha256"]


def test_explicit_body_override_beats_bare_heading_detection_without_exporting_marker(tmp_path):
    path = write_sample(tmp_path)
    snapshot = build_outline_snapshot(path)
    apply_outline_changes(path, snapshot["source_sha256"], [{"line_index": 6, "level": 0}])
    raw = path.read_text(encoding="utf-8")
    assert "<!--SXUPAPER:BODY-->1.2 遗漏标题" in raw
    parsed = parse_markdown(path)
    assert "1.2 遗漏标题" in "\n".join(section.content for section in parsed.sections)
    assert all(section.title != "1.2 遗漏标题" for section in parsed.sections)


def test_source_hash_conflict_and_protected_line_leave_file_untouched(tmp_path):
    path = write_sample(tmp_path)
    snapshot = build_outline_snapshot(path)
    before = path.read_bytes()
    with pytest.raises(OutlineEditError, match="已经变化") as stale:
        apply_outline_changes(path, "0" * 64, [{"line_index": 3, "level": 2}])
    assert stale.value.code == "outline_source_changed"
    with pytest.raises(OutlineEditError) as protected:
        apply_outline_changes(path, snapshot["source_sha256"], [{"line_index": 0, "level": 1}])
    assert protected.value.code == "outline_line_protected"
    assert path.read_bytes() == before


def test_level_gap_is_reported_with_exact_line_and_does_not_write(tmp_path):
    path = write_sample(tmp_path)
    snapshot = build_outline_snapshot(path)
    before = path.read_bytes()
    with pytest.raises(OutlineEditError) as error:
        apply_outline_changes(path, snapshot["source_sha256"], [{"line_index": 2, "level": 3}])
    assert error.value.code == "outline_level_gap"
    assert error.value.line_index == 2
    assert path.read_bytes() == before

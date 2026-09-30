from __future__ import annotations

import asyncio
from pathlib import Path

import pytest
from fastapi import HTTPException

from backend.api import routes

SAMPLE = """论文标题至少需要八个汉字

# 第一章 绪论
这是一段正文。
"""


def configure_runtime(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> Path:
    input_dir = tmp_path / "input"
    input_dir.mkdir()
    monkeypatch.setattr(routes, "PROJECT_ROOT", tmp_path)
    monkeypatch.setattr(routes, "INPUT_DIR", input_dir)
    path = input_dir / "paper.md"
    path.write_text(SAMPLE, encoding="utf-8", newline="\n")
    return path


def test_md_read_returns_hash_and_save_rejects_stale_source(monkeypatch, tmp_path):
    path = configure_runtime(monkeypatch, tmp_path)
    opened = asyncio.run(routes.api_get_md("input/paper.md"))
    assert opened["source_sha256"]

    path.write_text(SAMPLE + "外部修改。\n", encoding="utf-8", newline="\n")
    before = path.read_bytes()
    with pytest.raises(HTTPException) as conflict:
        asyncio.run(routes.api_save_md(routes.MdSaveRequest(
            path="input/paper.md",
            content=SAMPLE + "用户修改。\n",
            source_sha256=opened["source_sha256"],
        )))
    assert conflict.value.status_code == 409
    assert path.read_bytes() == before


def test_md_save_is_atomic_and_returns_next_hash(monkeypatch, tmp_path):
    path = configure_runtime(monkeypatch, tmp_path)
    opened = asyncio.run(routes.api_get_md("input/paper.md"))
    updated = SAMPLE.replace("这是一段正文。", "这是用户确认后的正文。")
    result = asyncio.run(routes.api_save_md(routes.MdSaveRequest(
        path="input/paper.md", content=updated, source_sha256=opened["source_sha256"]
    )))
    assert result["success"] is True
    assert result["source_sha256"] != opened["source_sha256"]
    assert path.read_text(encoding="utf-8") == updated
    assert not list(path.parent.glob(".paper.edit-*.md"))

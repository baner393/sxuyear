"""Safe, line-addressed chapter-structure editing for imported Markdown."""

from __future__ import annotations

import hashlib
import os
import re
import tempfile
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable

from .parser import (
    BODY_OVERRIDE_PREFIX,
    ParsedDocument,
    _detect_level,
    _is_known,
    _parse_metadata_block,
    parse_markdown,
)


class OutlineEditError(ValueError):
    """A user-correctable outline edit error with a stable code."""

    def __init__(self, code: str, message: str, *, line_index: int | None = None):
        super().__init__(message)
        self.code = code
        self.line_index = line_index


@dataclass(frozen=True)
class TextDocument:
    raw: bytes
    text: str
    encoding: str
    lines: list[str]


def _read_document(path: Path) -> TextDocument:
    raw = path.read_bytes()
    encoding = "utf-8-sig" if raw.startswith(b"\xef\xbb\xbf") else "utf-8"
    text = raw.decode(encoding)
    return TextDocument(raw=raw, text=text, encoding=encoding, lines=text.splitlines(keepends=True))


def _line_body(line: str) -> tuple[str, str]:
    body = line.rstrip("\r\n")
    return body, line[len(body):]


def _display_text(line: str) -> str:
    body, _ = _line_body(line)
    if body.startswith(BODY_OVERRIDE_PREFIX):
        body = body[len(BODY_OVERRIDE_PREFIX):]
    return re.sub(r"^#{1,6}\s*", "", body).strip()


def _level(line: str) -> int:
    body, _ = _line_body(line)
    stripped = body.strip()
    if body.startswith(BODY_OVERRIDE_PREFIX):
        return 0
    explicit = re.match(r"^(#{1,6})\s+", stripped)
    if explicit:
        return min(len(explicit.group(1)), 3)
    if _is_known(stripped):
        return 1
    return min(_detect_level(stripped), 3)


def _unmarkable_reason(line: str, index: int, metadata_end: int) -> str | None:
    text = _display_text(line)
    stripped, _ = _line_body(line)
    stripped = stripped.strip()
    if index < metadata_end:
        return "metadata"
    if not text:
        return "empty"
    if stripped.startswith("|"):
        return "table"
    if stripped.startswith("!["):
        return "image"
    if re.match(r"^\[\^\d+\]:", stripped):
        return "footnote"
    if re.fullmatch(r"(?:<!--FORMULA:\d+-->)+", stripped):
        return "formula"
    if stripped.startswith("<!--TABLE:"):
        return "table"
    return None


def _metadata_end(lines: list[str]) -> int:
    plain = [_line_body(line)[0] for line in lines]
    return _parse_metadata_block(plain, ParsedDocument())


def _nearest_context(lines: list[str], index: int, direction: int) -> str:
    cursor = index + direction
    while 0 <= cursor < len(lines):
        text = _display_text(lines[cursor])
        if text:
            return text[:160]
        cursor += direction
    return ""


def build_outline_snapshot(path: Path) -> dict:
    document = _read_document(path)
    metadata_end = _metadata_end(document.lines)
    rows = []
    headings = []
    for index, line in enumerate(document.lines):
        level = _level(line)
        reason = _unmarkable_reason(line, index, metadata_end)
        text = _display_text(line)
        row = {
            "line_index": index,
            "text": text,
            "level": level,
            "is_heading": level > 0,
            "markable": reason is None,
            "disabled_reason": reason,
            "forced_body": _line_body(line)[0].startswith(BODY_OVERRIDE_PREFIX),
            "before": _nearest_context(document.lines, index, -1),
            "after": _nearest_context(document.lines, index, 1),
        }
        rows.append(row)
        if level > 0:
            headings.append(dict(row))
    return {
        "source_sha256": hashlib.sha256(document.raw).hexdigest(),
        "line_count": len(document.lines),
        "headings": headings,
        "lines": rows,
    }


def _set_level(line: str, level: int) -> str:
    body, ending = _line_body(line)
    forced = body.startswith(BODY_OVERRIDE_PREFIX)
    if forced:
        body = body[len(BODY_OVERRIDE_PREFIX):]
    text = re.sub(r"^#{1,6}\s*", "", body).strip()
    if level == 0:
        # Bare numbered headings and known section names would otherwise be immediately
        # re-detected. Always use the parser-owned override for an explicit body choice.
        return f"{BODY_OVERRIDE_PREFIX}{text}{ending}"
    return f"{'#' * level} {text}{ending}"


def _table_signature(document: ParsedDocument) -> list[tuple]:
    return [
        (table.caption, tuple(table.headers), tuple(tuple(row) for row in table.rows))
        for section in document.sections
        for table in section.tables
    ]


def _protected_tokens(text: str) -> dict[str, tuple[str, ...]]:
    patterns = {
        "images": r"!\[[^\]]*\]\([^\)]*\)",
        "formulas": r"<!--FORMULA:\d+-->",
        "footnotes": r"\[\^\d+\]",
        "citations": r"\[\d+(?:\s*[-,，、]\s*\d+)*\]",
    }
    return {name: tuple(re.findall(pattern, text)) for name, pattern in patterns.items()}


def _validate_hierarchy(lines: list[str]) -> None:
    seen: set[int] = set()
    metadata_end = _metadata_end(lines)
    for index, line in enumerate(lines):
        if _unmarkable_reason(line, index, metadata_end) is not None:
            continue
        level = _level(line)
        if level <= 0:
            continue
        if level > 1 and level - 1 not in seen:
            raise OutlineEditError(
                "outline_level_gap",
                f"“{_display_text(line)}”是{level}级标题，但上方没有{level - 1}级标题",
                line_index=index,
            )
        seen = {item for item in seen if item < level}
        seen.add(level)


def apply_outline_changes(
    path: Path, source_sha256: str, changes: Iterable[dict[str, int]]
) -> dict:
    before = _read_document(path)
    actual_hash = hashlib.sha256(before.raw).hexdigest()
    if actual_hash != source_sha256:
        raise OutlineEditError("outline_source_changed", "论文内容已经变化，请刷新后重新修改")

    metadata_end = _metadata_end(before.lines)
    normalized: dict[int, int] = {}
    for change in changes:
        try:
            index = int(change["line_index"])
            level = int(change["level"])
        except (KeyError, TypeError, ValueError) as exc:
            raise OutlineEditError("outline_change_invalid", "章节修改数据不完整") from exc
        if index in normalized:
            raise OutlineEditError("outline_change_duplicate", "同一段落出现了重复修改")
        if not 0 <= index < len(before.lines) or level not in (0, 1, 2, 3):
            raise OutlineEditError("outline_change_invalid", "章节修改超出允许范围")
        reason = _unmarkable_reason(before.lines[index], index, metadata_end)
        if reason is not None:
            raise OutlineEditError(
                "outline_line_protected", "该行属于受保护内容，不能标记为章节", line_index=index
            )
        normalized[index] = level
    if not normalized:
        raise OutlineEditError("outline_changes_empty", "没有需要保存的章节修改")

    after_lines = list(before.lines)
    for index, level in normalized.items():
        after_lines[index] = _set_level(before.lines[index], level)
    if len(after_lines) != len(before.lines):
        raise OutlineEditError("outline_line_count_changed", "保存会改变原文行数，已停止")
    for index, line in enumerate(before.lines):
        if index not in normalized and after_lines[index] != line:
            raise OutlineEditError("outline_unselected_changed", "未选择的正文发生变化，已停止")

    _validate_hierarchy(after_lines)
    after_text = "".join(after_lines)
    suffix = path.suffix or ".md"
    fd, temporary_name = tempfile.mkstemp(prefix=f".{path.stem}.outline-", suffix=suffix, dir=path.parent)
    os.close(fd)
    temporary = Path(temporary_name)
    try:
        encoded = after_text.encode(before.encoding)
        temporary.write_bytes(encoded)
        original_doc = parse_markdown(path)
        updated_doc = parse_markdown(temporary)
        if original_doc.metadata != updated_doc.metadata:
            raise OutlineEditError("outline_metadata_changed", "封面字段发生变化，已停止保存")
        if original_doc.footnotes != updated_doc.footnotes:
            raise OutlineEditError("outline_footnotes_changed", "脚注结构发生变化，已停止保存")
        if _table_signature(original_doc) != _table_signature(updated_doc):
            raise OutlineEditError("outline_tables_changed", "表格结构发生变化，已停止保存")
        if _protected_tokens(before.text) != _protected_tokens(after_text):
            raise OutlineEditError("outline_tokens_changed", "图片、公式或引用标记发生变化，已停止保存")
        os.replace(temporary, path)
    finally:
        temporary.unlink(missing_ok=True)
    return build_outline_snapshot(path)

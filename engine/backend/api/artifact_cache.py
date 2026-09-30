"""转换产物的内容寻址缓存。

键 = sha256(管线源码指纹 + md 字节 + 模板 docx 字节 + 模板档案 json 字节
          + 请求 config 规范化 JSON + photo/formula 资产摘要 + 模式)。
任一输入变化（包括改了转换器代码本身）键即变化，永不误命中；
语义相同但表示不同的 config 只会误判为不命中——多算一次，不会算错。

条目 = output/_cache/<key前16位>/ 下的 doc.docx / doc.pdf / meta.json，
meta.json 最后写入，它的存在即条目有效（写一半的条目不会被读到）。
mode="fast"（预览快速通道）与 "full"（成稿）是不同的产物，分开成键。

单用户本机应用，不做并发锁；LRU 按 last_used 清理。
"""

from __future__ import annotations

import hashlib
import json
import shutil
import time
from pathlib import Path
from typing import Any, Optional

_THESIS_DIR = Path(__file__).resolve().parent.parent / "thesis_builder"
_SCRIPTS_DIR = Path(__file__).resolve().parent.parent / "scripts"
_API_DIR = Path(__file__).resolve().parent

# 会影响产物字节的全部源码；改任何一个，全部缓存自动失效
_FINGERPRINT_FILES = [
    _THESIS_DIR / "converter.py",
    _THESIS_DIR / "postprocess.py",
    _THESIS_DIR / "styles.py",
    _THESIS_DIR / "template_merge.py",
    _THESIS_DIR / "parser.py",
    _SCRIPTS_DIR / "word_pass.ps1",
    _SCRIPTS_DIR / "docx2pdf.ps1",
    _SCRIPTS_DIR / "word_preview_worker.ps1",
    _API_DIR / "word_preview_worker.py",
]

_fingerprint: Optional[str] = None


def _pipeline_fingerprint() -> str:
    """启动后首次调用时对管线源码取哈希（约 300KB，<5ms），进程内缓存。"""
    global _fingerprint
    if _fingerprint is None:
        h = hashlib.sha256()
        for f in _FINGERPRINT_FILES:
            h.update(f.name.encode())
            h.update(f.read_bytes() if f.exists() else b"<missing>")
        _fingerprint = h.hexdigest()
    return _fingerprint


def _assets_digest(input_dir: Path) -> str:
    """photo/ 与 formula/ 的 (相对路径, 大小, mtime_ns) 摘要。

    转换器在构建时直接从磁盘读这两个目录（photo 的图片、formula 的
    OMML），漏掉它们会导致换图后命中旧文档。
    """
    parts = []
    for sub in ("photo", "formula"):
        d = input_dir / sub
        if not d.is_dir():
            continue
        for f in sorted(d.iterdir()):
            if f.is_file():
                st = f.stat()
                parts.append(f"{sub}/{f.name}:{st.st_size}:{st.st_mtime_ns}")
    return "|".join(parts)


class ArtifactCache:
    def __init__(self, cache_dir: Path,
                 max_entries: int = 40,
                 max_bytes: int = 2 * 1024 ** 3) -> None:
        self.cache_dir = cache_dir
        self.max_entries = max_entries
        self.max_bytes = max_bytes
        cache_dir.mkdir(parents=True, exist_ok=True)

    # ── 键 ──
    def key(self, mode: str, md_path: Path, template_path: Path,
            config: Optional[dict]) -> str:
        h = hashlib.sha256()
        h.update(_pipeline_fingerprint().encode())
        h.update(b"|mode:" + mode.encode())
        h.update(b"|md:" + hashlib.sha256(Path(md_path).read_bytes()).digest())
        template_path = Path(template_path)
        h.update(b"|tpl:" + hashlib.sha256(template_path.read_bytes()).digest())
        profile = template_path.with_suffix(".json")
        if profile.exists():
            h.update(b"|profile:" + hashlib.sha256(profile.read_bytes()).digest())
        h.update(b"|cfg:" + json.dumps(
            config or {}, sort_keys=True, ensure_ascii=False,
            separators=(",", ":")).encode())
        h.update(b"|assets:" + _assets_digest(Path(md_path).parent).encode())
        return h.hexdigest()

    # ── 读 ──
    def get(self, key: str) -> Optional[dict]:
        """命中返回 meta dict（附 docx/pdf 的 Path 或 None），并触摸 last_used。"""
        entry = self.cache_dir / key[:16]
        meta_path = entry / "meta.json"
        if not meta_path.exists():
            return None
        try:
            meta = json.loads(meta_path.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            return None
        if meta.get("key") != key:      # 16 位前缀撞车（理论上），当不命中
            return None
        meta["docx"] = entry / "doc.docx" if (entry / "doc.docx").exists() else None
        meta["pdf"] = entry / "doc.pdf" if (entry / "doc.pdf").exists() else None
        meta["last_used"] = time.time()
        try:
            meta_path.write_text(json.dumps(
                {k: v for k, v in meta.items() if k not in ("docx", "pdf")},
                ensure_ascii=False), encoding="utf-8")
        except OSError:
            pass
        return meta

    # ── 写（upsert：convert 先存 docx，之后 export 补 pdf）──
    def put(self, key: str, *, title: str,
            docx: Optional[Path] = None, pdf: Optional[Path] = None,
            pages: Optional[int] = None,
            warnings: Optional[list] = None,
            extra: Optional[dict[str, Any]] = None) -> None:
        entry = self.cache_dir / key[:16]
        entry.mkdir(parents=True, exist_ok=True)
        if docx is not None:
            shutil.copy2(docx, entry / "doc.docx")
        if pdf is not None:
            shutil.copy2(pdf, entry / "doc.pdf")
        old = {}
        meta_path = entry / "meta.json"
        if meta_path.exists():
            try:
                old = json.loads(meta_path.read_text(encoding="utf-8"))
            except (OSError, json.JSONDecodeError):
                old = {}
        now = time.time()
        meta = {
            "key": key,
            "title": title,
            "pages": pages if pages is not None else old.get("pages"),
            "warnings": warnings if warnings is not None else old.get("warnings", []),
            "created_at": old.get("created_at", now),
            "last_used": now,
        }
        # A source DOCX needs the section-parity plan when it is later turned
        # into a deliverable. Preserve arbitrary future cache metadata too.
        for key_, value in old.items():
            if key_ not in meta:
                meta[key_] = value
        if extra:
            meta.update(extra)
        meta_path.write_text(json.dumps(meta, ensure_ascii=False), encoding="utf-8")
        self.prune()

    # ── LRU 清理 ──
    def prune(self) -> None:
        entries = []
        for d in self.cache_dir.iterdir():
            meta_path = d / "meta.json"
            if not d.is_dir() or not meta_path.exists():
                continue
            try:
                meta = json.loads(meta_path.read_text(encoding="utf-8"))
            except (OSError, json.JSONDecodeError):
                shutil.rmtree(d, ignore_errors=True)
                continue
            size = sum(f.stat().st_size for f in d.iterdir() if f.is_file())
            entries.append((meta.get("last_used", 0), size, d))
        entries.sort()                              # last_used 最旧在前
        total = sum(e[1] for e in entries)
        while entries and (len(entries) > self.max_entries
                           or total > self.max_bytes):
            _, size, d = entries.pop(0)
            shutil.rmtree(d, ignore_errors=True)
            total -= size

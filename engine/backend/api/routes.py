"""API 路由定义。

主要端点（完整清单以本文件 @router 装饰器为准）：
- POST /api/convert     生成成稿 docx（缓存命中秒回）
- POST /api/preview     生成预览 PDF（fast 通道 + 缓存）
- POST /api/export-pdf  成稿 docx + PDF
- GET  /api/pdf/{title} 取 PDF 文件
- GET  /api/templates / /api/inputs / /api/report 等
"""

from __future__ import annotations

import asyncio
import hashlib
import json
import os
import re
import shutil
import subprocess
import tempfile
from pathlib import Path
from typing import Any, Dict, List, Optional

from fastapi import APIRouter, File, HTTPException, UploadFile
from fastapi.responses import FileResponse
from pydantic import BaseModel
from starlette.concurrency import run_in_threadpool

from ..process_utils import hidden_process_kwargs
from ..thesis_builder.converter import convert, finalize_docx
from ..thesis_builder.docx_import import (
    DocxImportError,
    extract_docx_to_md,
    validate_docx_bytes,
)
from ..thesis_builder.metadata_edit import update_metadata_block
from ..thesis_builder.outline_edit import (
    OutlineEditError,
    apply_outline_changes,
    build_outline_snapshot,
)
from ..thesis_builder.parser import parse_markdown
from ..thesis_builder.styles import ThesisConfig
from ..thesis_builder.template_merge import load_template_profile
from .artifact_cache import ArtifactCache
from .word_preview_worker import WORD_PREVIEW_WORKER

# ═══════════════════════════════════════════
# 路由器
# ═══════════════════════════════════════════

router = APIRouter()

# 项目代码目录与桌面用户数据目录分离；未设置环境变量时保留原版行为。
SOURCE_ROOT = Path(__file__).resolve().parent.parent.parent
PROJECT_ROOT = Path(os.environ.get("SXUPAPER_RUNTIME_ROOT", SOURCE_ROOT)).resolve()
TEMPLATE_DIR = Path(os.environ.get("SXUPAPER_TEMPLATE_DIR", PROJECT_ROOT / "template")).resolve()
INPUT_DIR = PROJECT_ROOT / "input"
OUTPUT_DIR = PROJECT_ROOT / "output"
CONFIG_DIR = PROJECT_ROOT / "config"
REQUIRE_WORD = os.environ.get("SXUPAPER_REQUIRE_WORD", "").strip() == "1"
TEMPLATE_PINNING_REQUIRED = os.environ.get("SXUPAPER_ONLINE_REQUIRED", "").strip() == "1"

# 预览产物（docx/pdf）隔离目录：不按「导出」时 output/ 根目录保持干净
PREVIEW_DIR = OUTPUT_DIR / "_preview"

# 确保目录存在
for directory in (INPUT_DIR, OUTPUT_DIR, PREVIEW_DIR, CONFIG_DIR, TEMPLATE_DIR):
    directory.mkdir(parents=True, exist_ok=True)


def _pdf_failure_detail() -> str:
    if REQUIRE_WORD:
        return "PDF 转换失败：桌面版需要安装并正常启动 Microsoft Word，不会改用 WPS 或 LibreOffice"
    return "PDF 转换失败，请确保 Microsoft Word 或 LibreOffice 已安装"


def _workspace_output_dir(input_path: Path) -> Path:
    """Keep deliverables from different uploaded papers in separate directories."""
    try:
        relative = input_path.resolve().relative_to(INPUT_DIR.resolve())
    except ValueError:
        # Test doubles and development callers may inject a validated path
        # directly; production API paths are constrained by _resolve_input.
        relative = Path(input_path.name)
    if len(relative.parts) > 1 and relative.parts[0].isdigit():
        workspace_id = relative.parts[0]
    else:
        workspace_id = hashlib.sha256(relative.as_posix().encode("utf-8")).hexdigest()[:16]
    destination = OUTPUT_DIR / "workspaces" / workspace_id
    destination.mkdir(parents=True, exist_ok=True)
    return destination


def _enforce_template_pin(input_path: Path, template_path: Path, confirmed: bool) -> None:
    """Pin one immutable template hash per paper until the user confirms a change."""
    if not TEMPLATE_PINNING_REQUIRED:
        return
    output_dir = _workspace_output_dir(input_path)
    pin_path = output_dir / ".template-pin.json"
    requested = {
        "path": template_path.resolve().relative_to(PROJECT_ROOT).as_posix(),
        "sha256": hashlib.sha256(template_path.read_bytes()).hexdigest(),
    }
    current = None
    try:
        current = json.loads(pin_path.read_text(encoding="utf-8"))
    except (FileNotFoundError, OSError, ValueError, TypeError):
        pass
    if current and current.get("sha256") != requested["sha256"] and not confirmed:
        raise HTTPException(
            status_code=409,
            detail={
                "code": "template_change_confirmation_required",
                "current": current,
                "requested": requested,
            },
        )
    if current != requested:
        temporary = pin_path.with_suffix(".tmp")
        temporary.write_text(json.dumps(requested, ensure_ascii=False, indent=2), encoding="utf-8")
        temporary.replace(pin_path)


# 转换产物缓存：键含管线源码指纹，改代码自动全量失效
ARTIFACT_CACHE = ArtifactCache(OUTPUT_DIR / "_cache")

# 单槽 Word 任务锁：Word COM 单实例不可并发（HANDOFF §3.1），转换类请求
# 在此排队；重活丢进线程池跑，事件循环保持空转——转换进行中，其它端点
# （含缓存命中的转换请求）照常秒回，浏览器标签不再卡死。
# 同键 miss 在取得锁后会二次查缓存，多个相同请求只会实际渲染一次。
WORD_SLOT = asyncio.Lock()


def _restore_or_build_source(
    source_key: str, input_path: Path, template_path: Path, config: Optional[dict], output_dir: Path
) -> dict:
    """Return a pre-pagination DOCX, copied into ``output_dir``.

    A fast preview and a deliverable export start from byte-for-byte the same
    Python-generated document.  The source cache shares that work, while the
    final pagination pass always runs on a copy so it can never mutate the
    preview source (it writes TOC results and filler-page breaks in place).
    This helper is called under ``WORD_SLOT``: a cache miss will immediately
    continue into a Word render/finalization, so serializing prevents duplicate
    builds for identical requests without adding another lock.
    """
    hit = ARTIFACT_CACHE.get(source_key)
    if hit and hit["docx"] is not None and isinstance(hit.get("section_parities"), list):
        path = output_dir / f"{hit['title']}.docx"
        shutil.copy2(hit["docx"], path)
        return {
            "output_path": str(path),
            "title": hit["title"],
            "pages": hit.get("pages"),
            "warnings": list(hit.get("warnings", [])),
            "section_parities": hit["section_parities"],
        }

    result = convert(
        input_md=str(input_path),
        template_path=str(template_path),
        config=config,
        output_dir=str(output_dir),
        fast=True,
    )
    docx_path = Path(result["output_path"])
    ARTIFACT_CACHE.put(
        source_key,
        title=docx_path.stem,
        docx=docx_path,
        pages=result["pages"],
        warnings=result["warnings"],
        extra={"section_parities": result["section_parities"]},
    )
    result["title"] = docx_path.stem
    return result


def _finalize_source(source: dict) -> dict:
    """Apply the established Word parity/TOC pass to a copied source DOCX."""
    warnings = list(source.get("warnings", []))
    pages = finalize_docx(
        Path(source["output_path"]),
        source["section_parities"],
        warnings,
        fields_already_updated=source.get("fields_already_updated", False),
    )
    return {
        "success": True,
        "output_path": source["output_path"],
        "preview_url": f"/api/preview/{source['title']}",
        "pages": pages or source.get("pages") or 1,
        "warnings": warnings,
    }


# ═══════════════════════════════════════════
# 请求/响应模型
# ═══════════════════════════════════════════


class ConvertRequest(BaseModel):
    """转换请求"""

    input_md: str
    template_path: str = "template/山财学年论文模板.docx"
    config: Optional[Dict[str, Any]] = None
    confirm_template_change: bool = False


class ConvertResponse(BaseModel):
    """转换响应"""

    success: bool
    output_path: str
    preview_url: str
    pages: int
    warnings: List[str]


class PreviewRequest(BaseModel):
    """预览请求（与转换请求相同）"""

    input_md: str
    template_path: str = "template/山财学年论文模板.docx"
    config: Optional[Dict[str, Any]] = None
    confirm_template_change: bool = False
    accurate_pagination: bool = False


class PreviewResponse(BaseModel):
    """预览响应"""

    success: bool
    pdf_url: str
    total_pages: int
    warnings: List[str] = []


class ConfigSaveRequest(BaseModel):
    """配置保存请求"""

    name: str
    config: Dict[str, Any]


class ReportRequest(BaseModel):
    """导入检查与修正请求"""

    input_md: str


class OutlineItem(BaseModel):
    """导入检查与修正里的章节条目"""

    level: int
    title: str
    type: str
    tables: int = 0
    chars: int = 0


class ReportResponse(BaseModel):
    """导入检查与修正：转换前让用户确认解析结果"""

    success: bool
    md_path: str  # 中间 md 相对路径（编辑目标）
    source: str  # 'docx' | 'md'
    has_docx: bool = False  # 存在原始 Word 文件（语料库收录需要它）
    metadata: Dict[str, str]
    missing_fields: List[str]  # 标准 10 字段里缺失的
    outline: List[OutlineItem]
    stats: Dict[str, int]  # footnotes / tables / images / formulas
    warnings: List[str] = []


class CorpusAddRequest(BaseModel):
    """语料库收录请求"""

    path: str  # 输入路径（md 或 docx，须在 input/ 下）
    note: str = ""  # 一句话：这份文档哪里识别得不对


class MdSaveRequest(BaseModel):
    """中间 md 保存请求"""

    path: str
    content: str
    source_sha256: str


class MetadataSaveRequest(BaseModel):
    """封面元数据表单保存请求"""

    path: str  # 中间 md 相对路径
    metadata: Dict[str, str]  # 用户填写的字段（留空 = 跳过该栏）


class OutlineChange(BaseModel):
    line_index: int
    level: int


class OutlineSaveRequest(BaseModel):
    path: str
    source_sha256: str
    changes: List[OutlineChange]


# ═══════════════════════════════════════════
# 转换端点
# ═══════════════════════════════════════════


@router.post("/convert", response_model=ConvertResponse)
async def api_convert(req: ConvertRequest):
    """将 Markdown（或 Word 文档，自动先提取）转换为 DOCX"""
    input_path = _resolve_input(req.input_md)
    template_path = _resolve_template(req.template_path)

    if not input_path.exists():
        raise HTTPException(status_code=404, detail=f"输入文件不存在: {req.input_md}")
    if not template_path.exists():
        raise HTTPException(status_code=404, detail=f"模板文件不存在: {req.template_path}")

    _enforce_template_pin(input_path, template_path, req.confirm_template_change)

    input_path, extract_warnings = await _ensure_md_input(input_path)
    output_dir = _workspace_output_dir(input_path)

    try:
        cache_key = ARTIFACT_CACHE.key("full", input_path, template_path, req.config)
        hit = ARTIFACT_CACHE.get(cache_key)
        if hit and hit["docx"] is not None and hit.get("pages"):
            dst = output_dir / f"{hit['title']}.docx"
            shutil.copy2(hit["docx"], dst)
            return ConvertResponse(
                success=True,
                output_path=str(dst),
                preview_url=f"/api/preview/{hit['title']}",
                pages=hit["pages"],
                warnings=extract_warnings + hit.get("warnings", []),
            )

        source_key = ARTIFACT_CACHE.key("source", input_path, template_path, req.config)
        async with WORD_SLOT:
            source = await run_in_threadpool(
                _restore_or_build_source,
                source_key,
                input_path,
                template_path,
                req.config,
                output_dir,
            )
            # A fresh fast-preview cache may contain a separately saved DOCX
            # whose fields were updated by Word. Reuse it only when it proves
            # it came from this exact source key and retains the same section
            # plan. The immutable source cache is never modified.
            fast_key = ARTIFACT_CACHE.key("fast", input_path, template_path, req.config)
            fast_hit = ARTIFACT_CACHE.get(fast_key)
            if (
                fast_hit
                and fast_hit["docx"] is not None
                and fast_hit.get("preview_fields_persisted") is True
                and fast_hit.get("source_key") == source_key
                and fast_hit.get("section_parities") == source["section_parities"]
            ):
                shutil.copy2(fast_hit["docx"], source["output_path"])
                source["fields_already_updated"] = True
            result = await run_in_threadpool(_finalize_source, source)
        docx_path = Path(result["output_path"])
        ARTIFACT_CACHE.put(
            cache_key,
            title=docx_path.stem,
            docx=docx_path,
            pages=result["pages"],
            warnings=result["warnings"],
        )
        result["warnings"] = extract_warnings + result["warnings"]
        return ConvertResponse(**result)
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"转换失败: {str(e)}")


# ═══════════════════════════════════════════
# 预览端点
# ═══════════════════════════════════════════


@router.post("/preview", response_model=PreviewResponse)
async def api_preview(req: PreviewRequest):
    """生成 PDF 预览（输入为 Word 文档时自动先提取）"""
    input_path = _resolve_input(req.input_md)
    template_path = _resolve_template(req.template_path)

    if not input_path.exists():
        raise HTTPException(status_code=404, detail=f"输入文件不存在: {req.input_md}")
    if not template_path.exists():
        raise HTTPException(status_code=404, detail=f"模板文件不存在: {req.template_path}")

    _enforce_template_pin(input_path, template_path, req.confirm_template_change)

    input_path, extract_warnings = await _ensure_md_input(input_path)

    try:
        preview_mode = "accurate" if req.accurate_pagination else "fast"
        cache_key = ARTIFACT_CACHE.key(preview_mode, input_path, template_path, req.config)
        hit = ARTIFACT_CACHE.get(cache_key)
        if hit and hit["pdf"] is not None:
            return PreviewResponse(
                success=True,
                pdf_url=f"/api/cache/{cache_key}/pdf",
                total_pages=hit.get("pages") or _count_pdf_pages(hit["pdf"]),
                warnings=extract_warnings + hit.get("warnings", []),
            )

        # 预览产物全部落在 _preview 隔离目录，不污染 output/ 根目录
        # fast=True：跳过 Word 分页后处理（零启动），奇偶空白页只在成稿里补；
        # 目录域由下面 PDF 导出会话内更新（update_fields=True），页码仍正确
        async with WORD_SLOT:
            # 同一份新预览可能被连点、或由多个标签页同时请求。第一个请求
            # 等待 Word 时，后面的请求会卡在锁上；拿到锁后必须再查一次，
            # 否则它会把已经完成的同一份文档又完整渲染一遍。
            hit = ARTIFACT_CACHE.get(cache_key)
            if hit and hit["pdf"] is not None:
                return PreviewResponse(
                    success=True,
                    pdf_url=f"/api/cache/{cache_key}/pdf",
                    total_pages=hit.get("pages") or _count_pdf_pages(hit["pdf"]),
                    warnings=extract_warnings + hit.get("warnings", []),
                )
            # Do not reuse a deterministic working directory. Word may still
            # hold a previous preview DOCX open across an app upgrade/restart;
            # a unique directory lets the new request proceed independently.
            PREVIEW_DIR.mkdir(parents=True, exist_ok=True)
            preview_dir = Path(
                tempfile.mkdtemp(prefix=f"{cache_key[:12]}-", dir=PREVIEW_DIR)
            )
            source_key = ARTIFACT_CACHE.key("source", input_path, template_path, req.config)
            result = await run_in_threadpool(
                _restore_or_build_source,
                source_key,
                input_path,
                template_path,
                req.config,
                preview_dir,
            )
            if req.accurate_pagination:
                # Rewritten text can change the physical page count.  Run the
                # authoritative parity/TOC pass before its preview so the back
                # cover and odd/even section placement match the deliverable.
                finalized = await run_in_threadpool(_finalize_source, result)
                result["pages"] = finalized["pages"]
                result["warnings"] = finalized["warnings"]
                result["fields_already_updated"] = True
            docx_path = Path(result["output_path"])
            title = docx_path.stem

            # 转换为 PDF
            pdf_warnings: List[str] = []
            conversion_details: dict = {}
            pdf_path = await run_in_threadpool(
                _convert_to_pdf,
                docx_path,
                title,
                preview_dir,
                update_fields=not req.accurate_pagination,
                preview=True,
                warnings=pdf_warnings,
                persist_preview_fields=not req.accurate_pagination,
                conversion_details=conversion_details,
            )
        if pdf_path is None:
            raise HTTPException(status_code=500, detail=_pdf_failure_detail())

        # 计算页数
        total_pages = _count_pdf_pages(pdf_path)
        ARTIFACT_CACHE.put(
            cache_key,
            title=title,
            docx=docx_path,
            pdf=pdf_path,
            pages=total_pages,
            warnings=result.get("warnings", []) + pdf_warnings,
            extra={
                "preview_fields_persisted": conversion_details.get(
                    "preview_fields_persisted", False
                ),
                "source_key": source_key,
                "section_parities": result["section_parities"],
                "accurate_pagination": req.accurate_pagination,
            },
        )

        return PreviewResponse(
            success=True,
            pdf_url=f"/api/cache/{cache_key}/pdf",
            total_pages=total_pages,
            warnings=extract_warnings + result.get("warnings", []) + pdf_warnings,
        )
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"预览生成失败: {str(e)}")


class ExportPdfResponse(BaseModel):
    """PDF 导出响应"""

    success: bool
    pdf_path: str
    total_pages: int
    warnings: List[str] = []


@router.post("/export-pdf", response_model=ExportPdfResponse)
async def api_export_pdf(req: PreviewRequest):
    """转换并把 PDF 保存到 output/ 目录（预览不落盘，只有此端点才导出）"""
    input_path = _resolve_input(req.input_md)
    template_path = _resolve_template(req.template_path)

    if not input_path.exists():
        raise HTTPException(status_code=404, detail=f"输入文件不存在: {req.input_md}")
    if not template_path.exists():
        raise HTTPException(status_code=404, detail=f"模板文件不存在: {req.template_path}")

    _enforce_template_pin(input_path, template_path, req.confirm_template_change)

    input_path, extract_warnings = await _ensure_md_input(input_path)
    output_dir = _workspace_output_dir(input_path)

    try:
        # 与 /api/convert 同键（同为成稿产物）：先转换后导出的场景可复用 docx
        cache_key = ARTIFACT_CACHE.key("full", input_path, template_path, req.config)
        hit = ARTIFACT_CACHE.get(cache_key)
        if hit and hit["docx"] is not None:
            title = hit["title"]
            docx_dst = output_dir / f"{title}.docx"
            shutil.copy2(hit["docx"], docx_dst)
            pdf_warnings: List[str] = []
            if hit["pdf"] is not None:
                pdf_path = output_dir / f"{title}.pdf"
                shutil.copy2(hit["pdf"], pdf_path)
            else:
                # 只缓存过 docx（先 convert 后 export）：补导 PDF 并回填缓存
                async with WORD_SLOT:
                    pdf_path = await run_in_threadpool(
                        _convert_to_pdf, docx_dst, title, output_dir, warnings=pdf_warnings
                    )
                if pdf_path is None:
                    raise HTTPException(status_code=500, detail=_pdf_failure_detail())
                ARTIFACT_CACHE.put(cache_key, title=title, pdf=pdf_path)
            return ExportPdfResponse(
                success=True,
                pdf_path=str(pdf_path.relative_to(PROJECT_ROOT)).replace("\\", "/"),
                total_pages=_count_pdf_pages(pdf_path),
                warnings=extract_warnings + hit.get("warnings", []) + pdf_warnings,
            )

        async with WORD_SLOT:
            source_key = ARTIFACT_CACHE.key("source", input_path, template_path, req.config)
            source = await run_in_threadpool(
                _restore_or_build_source,
                source_key,
                input_path,
                template_path,
                req.config,
                output_dir,
            )
            result = await run_in_threadpool(_finalize_source, source)
            docx_path = Path(result["output_path"])
            title = docx_path.stem

            # convert() 刚做完全量域更新并落盘，导出会话无需再更新（省 3-8s）
            pdf_warnings = []
            pdf_path = await run_in_threadpool(
                _convert_to_pdf, docx_path, title, output_dir, warnings=pdf_warnings
            )
        if pdf_path is None:
            raise HTTPException(status_code=500, detail=_pdf_failure_detail())

        ARTIFACT_CACHE.put(
            cache_key,
            title=title,
            docx=docx_path,
            pdf=pdf_path,
            pages=result["pages"],
            warnings=result.get("warnings", []),
        )

        return ExportPdfResponse(
            success=True,
            pdf_path=str(pdf_path.relative_to(PROJECT_ROOT)).replace("\\", "/"),
            total_pages=_count_pdf_pages(pdf_path),
            warnings=extract_warnings + result.get("warnings", []) + pdf_warnings,
        )
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"PDF 导出失败: {str(e)}")


@router.get("/cache/{key}/pdf")
async def api_get_cached_pdf(key: str):
    """直接提供内容寻址缓存中的预览 PDF，不复制回 _preview。"""
    if not re.fullmatch(r"[0-9a-f]{64}", key):
        raise HTTPException(status_code=404, detail="缓存不存在")
    hit = ARTIFACT_CACHE.get(key)
    if not hit or hit["pdf"] is None:
        raise HTTPException(status_code=404, detail="缓存不存在或已被清理")
    return FileResponse(
        str(hit["pdf"]), media_type="application/pdf", filename=f"{hit['title']}.pdf"
    )


@router.get("/pdf/{title}")
async def api_get_pdf(title: str, src: str = "output"):
    """获取 PDF 文件（src=preview 取预览隔离目录，默认取 output/ 导出目录）"""
    # FastAPI 自动 URL 解码 title 参数
    search_dir = PREVIEW_DIR if src == "preview" else OUTPUT_DIR
    for pdf_file in search_dir.rglob("*.pdf"):
        if pdf_file.stem == title:
            return FileResponse(
                str(pdf_file),
                media_type="application/pdf",
                filename=f"{title}.pdf",
            )

    raise HTTPException(status_code=404, detail=f"PDF 不存在: {title}")


# ═══════════════════════════════════════════
# 模板和输入列表
# ═══════════════════════════════════════════


@router.get("/templates")
async def api_list_templates():
    """列出可用的 DOCX 模板（有格式档案的用档案里的显示名）"""
    templates = []
    if TEMPLATE_DIR.exists():
        # Development keeps the frozen bundled fixtures. The signed-in desktop
        # product exposes only controller-distributed, hash-verified versions.
        bundled_files = [] if TEMPLATE_PINNING_REQUIRED else sorted(TEMPLATE_DIR.glob("*.docx"))
        for f in bundled_files:
            profile = load_template_profile(f)
            templates.append(
                {
                    "name": profile.get("name") or f.stem,
                    "path": f"template/{f.name}",
                    "size": f.stat().st_size,
                    "has_profile": bool(profile),
                }
            )
        index_path = TEMPLATE_DIR / "managed" / "index.json"
        try:
            managed = json.loads(index_path.read_text(encoding="utf-8")).get("templates", [])
        except (FileNotFoundError, OSError, ValueError, AttributeError):
            managed = []
        for item in managed:
            relative_path = item.get("path", "")
            try:
                template_path = _resolve_template(relative_path)
            except HTTPException:
                continue
            if not template_path.is_file():
                continue
            templates.append(
                {
                    "name": (
                        f"{item.get('school', '学校模板')}"
                        f"（{item.get('version', '未标版本')}）"
                    ),
                    "path": relative_path,
                    "size": template_path.stat().st_size,
                    "has_profile": template_path.with_suffix(".json").is_file(),
                    "managed": True,
                    "template_id": item.get("id"),
                    "school": item.get("school"),
                    "version": item.get("version"),
                }
            )
    return {"templates": templates}


@router.get("/template-profile")
async def api_template_profile(path: str):
    """获取模板的格式档案（选中模板时前端据此带出该校默认配置）"""
    template_path = _resolve_template(path)
    if not template_path.exists():
        raise HTTPException(status_code=404, detail=f"模板文件不存在: {path}")
    return load_template_profile(template_path)


_EXTRACT_WARN_FILE = "extract_warnings.json"


def _save_extract_warnings(md_path: Path, warnings: List[str]) -> None:
    """提取警告落盘（导入检查与修正随时可取，不只在上传瞬间可见）"""
    try:
        (md_path.parent / _EXTRACT_WARN_FILE).write_text(
            json.dumps(warnings, ensure_ascii=False, indent=1), encoding="utf-8"
        )
    except OSError:
        pass


def _load_extract_warnings(md_path: Path) -> List[str]:
    warn_file = md_path.parent / _EXTRACT_WARN_FILE
    if not warn_file.exists():
        return []
    try:
        return list(json.loads(warn_file.read_text(encoding="utf-8")))
    except (OSError, ValueError):
        return []


def _extracted_md_for(docx_file: Path) -> Optional[Path]:
    """docx 对应的已提取 md（同目录同名，或 <stem>/ 子目录内），无则 None"""
    for cand in (
        docx_file.with_suffix(".md"),
        docx_file.parent / docx_file.stem / f"{docx_file.stem}.md",
    ):
        if cand.exists():
            return cand
    return None


def _source_docx_for(md_path: Path) -> Optional[Path]:
    """md 对应的原始 docx（_extracted_md_for 的逆向），无则 None"""
    cand = md_path.with_suffix(".docx")
    if cand.exists():
        return cand
    if md_path.parent.name == md_path.stem:
        cand = md_path.parent.parent / f"{md_path.stem}.docx"
        if cand.exists():
            return cand
    return None


async def _ensure_md_input(input_path: Path) -> tuple[Path, List[str]]:
    """输入是 .docx 时先提取成 md（已提取且 docx 未更新则直接复用）"""
    if input_path.suffix.lower() != ".docx":
        return input_path, []
    existing = _extracted_md_for(input_path)
    if existing and existing.stat().st_mtime >= input_path.stat().st_mtime:
        return existing, []
    target_dir = existing.parent if existing else input_path.parent / input_path.stem
    target_dir.mkdir(parents=True, exist_ok=True)
    try:
        result = await run_in_threadpool(extract_docx_to_md, input_path, target_dir)
    except DocxImportError as exc:
        raise HTTPException(status_code=422, detail=str(exc))
    except Exception as exc:
        raise HTTPException(status_code=422, detail=f"Word 文档解析失败：{exc}")
    _save_extract_warnings(result.md_path, result.warnings)
    return result.md_path, result.warnings


@router.get("/inputs")
async def api_list_inputs():
    """列出可用的输入论文（md + 尚未提取/已更新的 Word 文档）"""
    inputs = []
    if INPUT_DIR.exists():
        for md_file in sorted(INPUT_DIR.rglob("*.md")):
            rel_path = md_file.relative_to(PROJECT_ROOT)
            inputs.append(
                {
                    "name": md_file.stem,
                    "path": str(rel_path).replace("\\", "/"),
                    "directory": str(md_file.parent.relative_to(PROJECT_ROOT)).replace("\\", "/"),
                }
            )
        # 直接放进 input/ 的 Word 文档：没提取过（或 docx 比 md 新）才列出，
        # 选中后 convert/preview 会自动提取
        for docx_file in sorted(INPUT_DIR.rglob("*.docx")):
            if docx_file.name.startswith("~$"):
                continue
            extracted = _extracted_md_for(docx_file)
            if extracted and extracted.stat().st_mtime >= docx_file.stat().st_mtime:
                continue
            rel_path = docx_file.relative_to(PROJECT_ROOT)
            inputs.append(
                {
                    "name": docx_file.name,
                    "path": str(rel_path).replace("\\", "/"),
                    "directory": str(docx_file.parent.relative_to(PROJECT_ROOT)).replace("\\", "/"),
                }
            )
    return {"inputs": inputs}


@router.post("/upload")
async def api_upload_file(file: UploadFile = File(...)):
    """上传 .md / .docx 文件到 input 目录（.docx 自动提取为 md）"""
    filename = file.filename or ""
    lower = filename.lower()
    if lower.endswith(".doc"):
        raise HTTPException(
            status_code=400, detail="不支持旧版 .doc 格式，请用 Word 打开后「另存为」.docx 再上传"
        )
    if not (lower.endswith(".md") or lower.endswith(".docx")):
        raise HTTPException(status_code=400, detail="仅支持 .md / .docx 文件")
    is_docx = lower.endswith(".docx")

    # 大小限制：论文 docx 常含大图，放宽到 30MB
    max_size = 30 * 1024 * 1024 if is_docx else 10 * 1024 * 1024

    # 清理文件名（防止路径遍历）
    safe_filename = Path(filename).name
    if not safe_filename or safe_filename.startswith("."):
        raise HTTPException(status_code=400, detail="无效的文件名")

    # 读取并检查大小
    content = await file.read()
    if len(content) > max_size:
        raise HTTPException(status_code=413, detail=f"文件过大，最大 {max_size // (1024 * 1024)}MB")

    if is_docx:
        try:
            validate_docx_bytes(content)
        except DocxImportError as exc:
            raise HTTPException(status_code=400, detail=str(exc))

    # 确定保存路径（自动编号子目录）
    existing_dirs = [d.name for d in INPUT_DIR.iterdir() if d.is_dir() and d.name.isdigit()]
    next_num = max((int(d) for d in existing_dirs), default=0) + 1
    upload_dir = INPUT_DIR / str(next_num)
    upload_dir.mkdir(parents=True, exist_ok=True)

    # 保存原始文件（docx 原件保留，便于溯源和重新提取）
    file_path = upload_dir / safe_filename
    file_path.write_bytes(content)

    warnings: List[str] = []
    if is_docx:
        # 提取是 CPU/IO 密集操作，丢线程池避免阻塞事件循环
        try:
            result = await run_in_threadpool(extract_docx_to_md, file_path, upload_dir)
        except DocxImportError as exc:
            shutil.rmtree(upload_dir, ignore_errors=True)
            raise HTTPException(status_code=422, detail=str(exc))
        except Exception as exc:
            shutil.rmtree(upload_dir, ignore_errors=True)
            raise HTTPException(status_code=422, detail=f"Word 文档解析失败：{exc}")
        file_path = result.md_path
        warnings = result.warnings
        _save_extract_warnings(result.md_path, warnings)

    # 返回文件信息（path 始终指向 md，前端与后续转换无感）
    rel_path = file_path.relative_to(PROJECT_ROOT)
    return {
        "success": True,
        "name": file_path.stem,
        "path": str(rel_path).replace("\\", "/"),
        "directory": str(upload_dir.relative_to(PROJECT_ROOT)).replace("\\", "/"),
        "warnings": warnings,
        "source": "docx" if is_docx else "md",
    }


# ═══════════════════════════════════════════
# 导入检查与修正 / 中间 md 编辑
# ═══════════════════════════════════════════

# 元数据标准字段（顺序 = 报告展示顺序）
_META_KEYS = [
    "title",
    "engTitle",
    "name",
    "studentId",
    "majorCode",
    "class",
    "major",
    "college",
    "advisor",
    "date",
]


@router.post("/report", response_model=ReportResponse)
async def api_report(req: ReportRequest):
    """导入检查与修正：解析（docx 先自动提取）后返回元数据、章节树、统计与提示，
    供用户在转换前确认识别是否正确。"""
    input_path = _resolve_input(req.input_md)
    if not input_path.exists():
        raise HTTPException(status_code=404, detail="输入文件不存在")
    md_path, extract_warnings = await _ensure_md_input(input_path)
    if not extract_warnings:
        extract_warnings = _load_extract_warnings(md_path)

    try:
        doc = await run_in_threadpool(parse_markdown, md_path)
        md_text = md_path.read_text(encoding="utf-8-sig")
    except Exception as exc:
        raise HTTPException(status_code=422, detail=f"解析失败：{exc}")

    metadata = {k: v for k, v in doc.metadata.items() if v}
    outline = [
        OutlineItem(
            level=sec.level,
            title=sec.title,
            type=sec.section_type,
            tables=len(sec.tables),
            chars=len(sec.content or ""),
        )
        for sec in doc.sections
    ]
    stats = {
        "footnotes": len(doc.footnotes),
        "tables": sum(len(sec.tables) for sec in doc.sections),
        "images": len(re.findall(r"!\[[^\]]*\]\(", md_text)),
        "formulas": len(set(re.findall(r"<!--FORMULA:(\d+)-->", md_text))),
    }
    is_docx = input_path.suffix.lower() == ".docx"
    return ReportResponse(
        success=True,
        md_path=str(md_path.relative_to(PROJECT_ROOT)).replace("\\", "/"),
        source="docx" if is_docx else "md",
        has_docx=is_docx or _source_docx_for(md_path) is not None,
        metadata=metadata,
        missing_fields=[k for k in _META_KEYS if not metadata.get(k)],
        outline=outline,
        stats=stats,
        warnings=extract_warnings,
    )


@router.get("/md")
async def api_get_md(path: str):
    """读取中间 md 原文（导入检查与修正的高级编辑入口）"""
    md_path = _resolve_input(path)
    if md_path.suffix.lower() != ".md" or not md_path.exists():
        raise HTTPException(status_code=404, detail="md 文件不存在")
    return {
        "path": str(md_path.relative_to(PROJECT_ROOT)).replace("\\", "/"),
        "content": md_path.read_text(encoding="utf-8-sig"),
        "source_sha256": hashlib.sha256(md_path.read_bytes()).hexdigest(),
    }


@router.get("/outline-editor")
async def api_outline_editor(path: str):
    """Return physical-line identities used by the staged chapter editor."""
    md_path = _resolve_input(path)
    if md_path.suffix.lower() != ".md" or not md_path.exists():
        raise HTTPException(status_code=404, detail="中间论文文件不存在")
    return build_outline_snapshot(md_path)


@router.post("/outline-editor/apply")
async def api_apply_outline(req: OutlineSaveRequest):
    """Atomically apply heading-only edits after hash and structure validation."""
    md_path = _resolve_input(req.path)
    if md_path.suffix.lower() != ".md" or not md_path.exists():
        raise HTTPException(status_code=404, detail="中间论文文件不存在")
    try:
        return apply_outline_changes(
            md_path,
            req.source_sha256,
            [change.model_dump() for change in req.changes],
        )
    except OutlineEditError as exc:
        raise HTTPException(
            status_code=409 if exc.code == "outline_source_changed" else 422,
            detail={"code": exc.code, "message": str(exc), "line_index": exc.line_index},
        ) from exc


@router.post("/metadata/save")
async def api_save_metadata(req: MetadataSaveRequest):
    """表单化保存封面元数据：逐字段按 parser 谓词校验后写回 md 元数据块。
    不合格的字段不写入（会污染槽位匹配），原因通过 rejected 返回给前端内联展示。"""
    md_path = _resolve_input(req.path)
    if md_path.suffix.lower() != ".md":
        raise HTTPException(status_code=400, detail="只能保存 .md 文件")
    if not md_path.is_relative_to(INPUT_DIR.resolve()):
        raise HTTPException(status_code=403, detail="只能修改 input 目录下的文件")
    if not md_path.exists():
        raise HTTPException(status_code=404, detail="md 文件不存在")
    md_text = md_path.read_text(encoding="utf-8-sig")
    new_text, applied, rejected = update_metadata_block(md_text, req.metadata)
    md_path.write_text(new_text, encoding="utf-8", newline="\n")
    return {"success": True, "applied": applied, "rejected": rejected}


@router.post("/corpus/add")
async def api_corpus_add(req: CorpusAddRequest):
    """把识别有问题的文档（自动脱敏后）收录进回归语料库。

    语料库 = backend/tests/corpus/，每个 case 是一份脱敏文档 + 期望提取结果，
    修复识别规则后由 pytest 保证同样的错不再犯。收录需要原始 Word 文件。"""
    input_path = _resolve_input(req.path)
    if not input_path.is_relative_to(INPUT_DIR.resolve()):
        raise HTTPException(status_code=403, detail="只能收录 input 目录下的文档")
    if input_path.suffix.lower() == ".docx":
        docx = input_path
    else:
        docx = _source_docx_for(input_path)
    if docx is None or not docx.exists():
        raise HTTPException(
            status_code=400, detail="这份输入没有对应的原始 Word 文件，无法收录（语料库只收 .docx）"
        )

    # 语料工具在 tests 包里（开发态工具），仅此处引用
    from ..tests.corpus_tool import add_case

    try:
        result = await run_in_threadpool(add_case, docx, None, req.note.strip() or None)
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"收录失败：{exc}")
    return {
        "success": True,
        "case": result["case"],
        "scrubbed": len(result["scrub"]),
        "residuals": result["residuals"],
    }


@router.post("/md/save")
async def api_save_md(req: MdSaveRequest):
    """保存用户修正后的中间 md（只允许改 input/ 下已存在的 md；
    保存后 mtime 更新，docx 不会覆盖用户修正）"""
    md_path = _resolve_input(req.path)
    if md_path.suffix.lower() != ".md":
        raise HTTPException(status_code=400, detail="只能保存 .md 文件")
    if not md_path.is_relative_to(INPUT_DIR.resolve()):
        raise HTTPException(status_code=403, detail="只能修改 input 目录下的文件")
    if not md_path.exists():
        raise HTTPException(status_code=404, detail="md 文件不存在")
    current_hash = hashlib.sha256(md_path.read_bytes()).hexdigest()
    if current_hash != req.source_sha256:
        raise HTTPException(
            status_code=409,
            detail={
                "code": "md_source_changed",
                "message": "论文内容已被其他操作修改，请刷新后重新编辑",
            },
        )

    descriptor, temporary_name = tempfile.mkstemp(
        prefix=f".{md_path.stem}.edit-", suffix=md_path.suffix, dir=md_path.parent
    )
    os.close(descriptor)
    temporary = Path(temporary_name)
    try:
        temporary.write_text(req.content, encoding="utf-8", newline="\n")
        try:
            parse_markdown(temporary)
        except Exception as exc:
            raise HTTPException(status_code=422, detail=f"原文结构校验失败：{exc}") from exc
        os.replace(temporary, md_path)
    finally:
        temporary.unlink(missing_ok=True)
    return {
        "success": True,
        "source_sha256": hashlib.sha256(md_path.read_bytes()).hexdigest(),
    }


# ═══════════════════════════════════════════
# 配置管理
# ═══════════════════════════════════════════


@router.get("/config/default")
async def api_get_default_config():
    """获取默认配置"""
    cfg = ThesisConfig()
    return cfg.to_dict()


@router.post("/config/save")
async def api_save_config(req: ConfigSaveRequest):
    """保存命名配置预设"""
    # 验证名称
    safe_name = re.sub(r"[^\w\-]", "_", req.name)
    if not safe_name:
        raise HTTPException(status_code=400, detail="无效的配置名称")

    config_file = CONFIG_DIR / f"{safe_name}.json"
    config_file.write_text(
        json.dumps(req.config, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    return {"success": True, "path": f"config/{safe_name}.json"}


@router.get("/config/list")
async def api_list_configs():
    """列出已保存的配置预设"""
    configs = []
    if CONFIG_DIR.exists():
        for f in sorted(CONFIG_DIR.glob("*.json")):
            configs.append(
                {
                    "name": f.stem,
                    "path": f"config/{f.name}",
                }
            )
    return {"configs": configs}


@router.get("/config/{name}")
async def api_load_config(name: str):
    """加载命名配置预设"""
    safe_name = re.sub(r"[^\w\-]", "_", name)
    config_file = CONFIG_DIR / f"{safe_name}.json"

    if not config_file.exists():
        raise HTTPException(status_code=404, detail=f"配置不存在: {name}")

    config = json.loads(config_file.read_text(encoding="utf-8"))
    return {"name": safe_name, "config": config}


# ═══════════════════════════════════════════
# 辅助函数
# ═══════════════════════════════════════════


def _validate_path(path: Path) -> Path:
    """验证路径在项目目录内（防止路径遍历攻击）"""
    try:
        resolved = path.resolve()
        project_resolved = PROJECT_ROOT.resolve()
        if not resolved.is_relative_to(project_resolved):
            raise HTTPException(status_code=403, detail="访问被拒绝")
        return resolved
    except (ValueError, OSError):
        raise HTTPException(status_code=400, detail="无效的路径")


def _resolve_input(input_md: str) -> Path:
    """解析输入文件路径"""
    path = Path(input_md)
    if not path.is_absolute():
        path = PROJECT_ROOT / path
    return _validate_path(path)


def _resolve_template(template_path: str) -> Path:
    """解析模板文件路径"""
    path = Path(template_path)
    if not path.is_absolute():
        path = PROJECT_ROOT / path
    return _validate_path(path)


def _find_libreoffice() -> Optional[str]:
    """查找 LibreOffice 可执行文件"""
    import shutil

    # Windows 路径
    candidates = [
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    ]
    for p in candidates:
        if Path(p).exists():
            return p

    # Linux/macOS
    for cmd in ("libreoffice", "soffice"):
        if shutil.which(cmd):
            return cmd

    return None


def _url_safe_title(title: str) -> str:
    """将标题转换为 URL 安全的字符串"""
    import urllib.parse

    return urllib.parse.quote(title, safe="")


def _convert_to_pdf(
    docx_path: Path,
    title: str,
    out_dir: Optional[Path] = None,
    update_fields: bool = False,
    preview: bool = False,
    warnings: Optional[List[str]] = None,
    persist_preview_fields: bool = False,
    conversion_details: Optional[dict] = None,
) -> Optional[Path]:
    """将 DOCX 转换为 PDF（优先 Word，后备 LibreOffice）。

    Args:
        update_fields: 让 Word 在导出会话内先重建域/目录（不回写 docx）。
            预览快速通道的 docx 跳过了分页后处理，必须传 True；
            完整 convert 刚更新并落盘过域，传 False 省 3-8s。
        preview: 只用于交互预览时，使用 Word 的屏幕优化 PDF。成稿和
            /export-pdf 一律保持原质量的 SaveAs2 输出。
        warnings: 传入列表时，回退 LibreOffice 等降级情况会追加人话警告。
    """
    # Word COM 按自身工作目录解析相对路径，必须转成绝对路径
    if not docx_path.is_absolute():
        docx_path = (PROJECT_ROOT / docx_path).resolve()
    if out_dir is None:
        out_dir = OUTPUT_DIR
    pdf_path = out_dir / f"{title}.pdf"
    # A title is stable across edits. Never let a failed render return an old
    # same-name PDF as if it represented the new source/configuration.
    try:
        pdf_path.unlink()
    except FileNotFoundError:
        pass

    # 方法 1：Microsoft Word（Windows）
    persisted = {"preview_fields_persisted": False}
    word_kwargs = {"update_fields": update_fields, "preview": preview}
    if persist_preview_fields:
        word_kwargs["persist_preview_fields"] = True
        word_kwargs["conversion_details"] = persisted
    if _convert_via_word(docx_path, pdf_path, **word_kwargs):
        if conversion_details is not None:
            conversion_details.update(persisted)
        return pdf_path

    if REQUIRE_WORD:
        return None

    # 方法 2：LibreOffice
    lo_cmd = _find_libreoffice()
    if lo_cmd:
        try:
            subprocess.run(
                [
                    lo_cmd,
                    "--headless",
                    "--convert-to",
                    "pdf",
                    "--outdir",
                    str(out_dir),
                    str(docx_path),
                ],
                capture_output=True,
                text=True,
                timeout=120,
                **hidden_process_kwargs(),
            )
            if pdf_path.exists():
                # --convert-to 不更新任何域：目录/页码停留在 docx 里的缓存值，
                # 预览快速通道的 docx 缓存还是占位符。绝不能无声吞掉这个降级
                if warnings is not None:
                    warnings.append(
                        "Word 不可用，本 PDF 由 LibreOffice 生成：目录与页码"
                        "未更新（可能过期或为占位符），排版仅供参考"
                    )
                return pdf_path
        except (subprocess.TimeoutExpired, FileNotFoundError):
            pass

    return None


def _convert_via_word(
    docx_path: Path,
    pdf_path: Path,
    update_fields: bool = False,
    preview: bool = False,
    *,
    persist_preview_fields: bool = False,
    conversion_details: Optional[dict] = None,
) -> bool:
    """使用 Microsoft Word 将 DOCX 转换为 PDF"""
    # Historical persistent Word work was reverted because it inserted page
    # breaks through a hot COM document. This worker is intentionally different:
    # it only opens an already-built preview, updates fields in memory, exports
    # and closes it. On any worker fault, use the proven one-shot script.
    if preview and update_fields:
        try:
            if persist_preview_fields:
                WORD_PREVIEW_WORKER.export(docx_path, pdf_path, persist_fields=True)
                if conversion_details is not None:
                    conversion_details["preview_fields_persisted"] = True
            else:
                WORD_PREVIEW_WORKER.export(docx_path, pdf_path)
            return True
        except Exception as exc:
            print(f"[word-preview-worker] fallback to one-shot Word: {exc}")
    elif not preview and not update_fields:
        try:
            WORD_PREVIEW_WORKER.export_final_pdf(docx_path, pdf_path)
            return True
        except Exception as exc:
            print(f"[word-final-pdf-worker] fallback to one-shot Word: {exc}")

    ps_script = Path(__file__).parent.parent / "scripts" / "docx2pdf.ps1"
    if not ps_script.exists():
        return False

    try:
        cmd = [
            "powershell.exe",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            str(ps_script),
            "-InputDocx",
            str(docx_path),
            "-OutputPdf",
            str(pdf_path),
        ]
        if update_fields:
            cmd.append("-UpdateFields")
        if preview:
            cmd.append("-Preview")
        result = subprocess.run(
            cmd,
            capture_output=True,
            timeout=120,
            **hidden_process_kwargs(),
        )
        ok = result.returncode == 0 and pdf_path.exists()
        if not ok:
            # 输出诊断信息，便于定位 Word 转换失败原因
            stderr = result.stderr.decode("gbk", errors="replace").strip()
            stdout = result.stdout.decode("gbk", errors="replace").strip()
            print(f"[docx2pdf] returncode={result.returncode} pdf_exists={pdf_path.exists()}")
            if stdout:
                print(f"[docx2pdf] stdout: {stdout[:500]}")
            if stderr:
                print(f"[docx2pdf] stderr: {stderr[:800]}")
        return ok
    except (subprocess.TimeoutExpired, FileNotFoundError) as e:
        print(f"[docx2pdf] 调用异常: {e}")
        return False


def _count_pdf_pages(pdf_path: Path) -> int:
    """计算 PDF 页数。

    用 pypdfium2 而非正则数 /Type /Page：Word 写经典 xref 表时正则碰巧
    是对的，但 LibreOffice 写压缩对象流，正则会数出 1。
    """
    try:
        import pypdfium2 as pdfium

        pdf = pdfium.PdfDocument(str(pdf_path))
        try:
            return max(1, len(pdf))
        finally:
            pdf.close()
    except Exception as e:
        print(f"PDF page count error: {e}")
        return 1

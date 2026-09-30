"""FastAPI 应用入口。

启动方式：
    cd backend
    uvicorn api.main:app --reload --host 0.0.0.0 --port 8000
"""

from __future__ import annotations

import os
import threading
from contextlib import asynccontextmanager
from pathlib import Path

from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles

from .routes import router
from .word_preview_worker import WORD_PREVIEW_WORKER

# ═══════════════════════════════════════════
# 应用初始化
# ═══════════════════════════════════════════


def warm_word_preview_worker() -> None:
    """Hide the first Word cold start behind backend startup when possible."""
    try:
        WORD_PREVIEW_WORKER.warm()
    except Exception as exc:
        # Preview/export will use the proven one-shot fallback on demand.
        print(f"[word-worker] startup warmup skipped: {exc}")


def start_word_preview_worker() -> None:
    if os.name == "nt" and os.environ.get("SXUPAPER_SKIP_WORD_WARMUP") != "1":
        threading.Thread(
            target=warm_word_preview_worker, name="word-preview-warmup", daemon=True
        ).start()


def close_word_preview_worker() -> None:
    """Close the app-owned Word process; never leave it running after stop."""
    WORD_PREVIEW_WORKER.close()


@asynccontextmanager
async def lifespan(_app: FastAPI):
    start_word_preview_worker()
    try:
        yield
    finally:
        close_word_preview_worker()


app = FastAPI(
    title="ThesisBuilder API",
    description="山西财经大学学年论文 Markdown→DOCX 自动排版工具",
    version="0.1.0",
    lifespan=lifespan,
)

# 开发态默认兼容原版前端；桌面态由主进程收窄到自己的随机回环地址。
configured_origins = os.environ.get("SXUPAPER_CORS_ORIGINS")
cors_origins = (
    [item.strip() for item in configured_origins.split(",") if item.strip()]
    if configured_origins
    else ["*"]
)
app.add_middleware(
    CORSMiddleware,
    allow_origins=cors_origins,
    allow_credentials=cors_origins != ["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

# 注册路由
app.include_router(router, prefix="/api")


# ═══════════════════════════════════════════
# 静态文件服务
# ═══════════════════════════════════════════

# 桌面版可把产物写到当前 Windows 用户目录；未设置时保持原版目录。
SOURCE_ROOT = Path(__file__).resolve().parent.parent.parent
PROJECT_ROOT = Path(os.environ.get("SXUPAPER_RUNTIME_ROOT", SOURCE_ROOT)).resolve()

# 输出目录（用于预览图片）
OUTPUT_DIR = PROJECT_ROOT / "output"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

# 挂载静态文件
app.mount("/output", StaticFiles(directory=str(OUTPUT_DIR)), name="output")


# ═══════════════════════════════════════════
# 健康检查
# ═══════════════════════════════════════════


@app.get("/")
async def root():
    """健康检查"""
    return {
        "service": "ThesisBuilder API",
        "version": "0.1.0",
        "status": "running",
    }


@app.get("/health")
async def health():
    """健康检查端点"""
    return {"status": "ok"}

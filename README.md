# SXUPaper 论文排版助手（开源版）

山西财经大学学年论文 Markdown → DOCX 自动排版引擎。

**体验完整功能请下载 Windows 客户端：** <https://sxupaper.b100.top/>

开源版包含论文排版引擎的核心算法、API 与回归测试；学校官方模板语料库、
AI 降重/内容优化、账号与套餐体系、卡密与支付、模板权益下发、桌面客户端与
在线控制平面均为**闭源商业核心**，不在本仓库内。

## 开源范围

- Markdown → 学校规范格式 DOCX 的排版转换（`engine/backend/thesis_builder`）
- DOCX 导入与识别（`docx_import`）、目录/大纲编辑（`outline_edit`）、
  元数据编辑、格式样式（`styles`）、模板合并（`template_merge`）
- 本地 API（`engine/backend/api`）：转换 / 预览 PDF / 导出 / 输入列表
- 回归测试与脱敏语料（`engine/backend/tests`）

## 闭源范围（不在本仓库，且不随本仓库分发）

- 学校官方模板语料库（docx 模板 + 各校格式档案）
- AI 降重 / 内容优化（LowerAI 引擎与云端服务）
- 账号 / VIP / 卡密 / 支付 / 额度权益（controller 控制平面）
- 模板权益下发与解锁、管理后台
- 桌面客户端（Electron 壳与在线客户端逻辑）

> 排版引擎需要用户自行提供学校模板 DOCX 才能产出最终文档。开源版不附带
> 任何学校模板文件；商用/学校授权模板通过官网客户端分发。

## 快速开始

```bash
# 后端（Python ≥ 3.10）
cd engine/backend
python -m venv .venv
.venv/Scripts/python.exe -m pip install -e ".[dev]"

# 运行 API（默认 127.0.0.1:8000）
.venv/Scripts/python.exe -m uvicorn api.main:app --reload --host 0.0.0.0 --port 8000

# 测试
.venv/Scripts/python.exe -m pytest tests -q
```

Windows 环境下最终 DOCX/PDF 渲染依赖本机 Microsoft Word（不静默回退
WPS/LibreOffice），预览/导出走 `engine/backend/scripts/docx2pdf.ps1`。

## API 概览

| 端点 | 说明 |
|---|---|
| `POST /api/convert` | Markdown → 成稿 DOCX（带缓存） |
| `POST /api/preview` | 生成预览 PDF（fast 通道 + 缓存） |
| `POST /api/export-pdf` | 成稿 DOCX + PDF |
| `GET  /api/pdf/{title}` | 取 PDF 文件 |
| `GET  /api/inputs` | 列出可用的输入论文 |
| `GET  /api/report` | 转换/预览报告 |

## 目录结构

```
engine/backend/
  api/               FastAPI 路由（convert/preview/export/inputs）
  thesis_builder/    排版引擎（解析、转换、导入、样式、模板合并）
  scripts/           Word 自动化（docx2pdf 等）
  tests/             回归测试 + 脱敏语料
deploy/
  promo/             官网宣传页（https://sxupaper.b100.top/）
  release-installer.ps1  发布四件套（exe / latest / sha256 / version）脚本
```

## 许可证

MIT License，见 [LICENSE](LICENSE)。本许可证仅覆盖本仓库代码；不授予对
闭源商业模块（学校模板语料库、AI 降重服务、账号/支付体系等）的任何权利。

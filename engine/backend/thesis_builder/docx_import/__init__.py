"""docx 导入：把用户上传的 Word 论文提取成本项目 md 契约格式。

用法（API 侧）:
    result = extract_docx_to_md(docx_path, output_dir)
    # output_dir 下产出 <stem>.md、photo/N.ext、formula/N.xml

调试 CLI:
    python -m backend.thesis_builder.docx_import <docx> <输出目录>
"""

from __future__ import annotations

import io
import re
import shutil
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import List

from .builder import build_markdown
from .reader import DocxReader


@dataclass
class ExtractResult:
    md_path: Path
    image_count: int = 0
    formula_count: int = 0
    warnings: List[str] = field(default_factory=list)


class DocxImportError(ValueError):
    """docx 无法导入（格式错误/无内容），message 面向用户"""


def validate_docx_bytes(content: bytes) -> None:
    """上传字节预检：旧版 .doc / 损坏 zip / 非 Word 包 → DocxImportError"""
    if content[:4] == b'\xd0\xcf\x11\xe0':
        raise DocxImportError("这是旧版 .doc 格式，请用 Word 打开后「另存为」.docx 再上传")
    try:
        with zipfile.ZipFile(io.BytesIO(content)) as zf:
            if 'word/document.xml' not in zf.namelist():
                raise DocxImportError("文件不是有效的 Word 文档（缺少正文部件）")
    except zipfile.BadZipFile:
        raise DocxImportError("文件已损坏或不是有效的 .docx 文档")


def extract_docx_to_md(docx_path: Path, output_dir: Path) -> ExtractResult:
    """主入口：解析 docx，写出 md + photo/ + formula/。

    Raises:
        DocxImportError: 文档无可提取文字（如扫描件）
    """
    docx_path = Path(docx_path)
    output_dir = Path(output_dir)
    reader = DocxReader(docx_path)
    try:
        fields = reader.read_cover_fields()
        blocks, anchors = reader.read_blocks()

        if not any(b.kind in ('heading', 'para') and b.text.strip() for b in blocks):
            raise DocxImportError("文档中没有可提取的文字内容（可能是扫描件或空文档）")

        # 重新提取时清掉上一轮的图片/公式，避免残留文件污染编号
        for sub in ('photo', 'formula'):
            shutil.rmtree(output_dir / sub, ignore_errors=True)

        # 只导出会被 md 引用区间内的图片（跳过封面校徽/评定表尾件里的图）
        start = anchors.get('body_start', 0)
        end = anchors.get('attachment', len(blocks))
        image_names = reader.export_images(blocks[start:end], output_dir / 'photo')
        formula_count = reader.export_formulas(output_dir / 'formula')

        warnings: List[str] = list(reader.warnings)
        md_text = build_markdown(reader, blocks, anchors, fields, image_names, warnings)

        # 清理导出后未被 md 引用的图片（如致谢后的校徽装饰图）
        used = set(re.findall(r'\]\((photo/[^)]+)\)', md_text))
        for rid, name in list(image_names.items()):
            if name not in used:
                (output_dir / name).unlink(missing_ok=True)
                del image_names[rid]

        md_path = output_dir / f"{docx_path.stem}.md"
        md_path.write_text(md_text, encoding='utf-8', newline='\n')
        return ExtractResult(
            md_path=md_path,
            image_count=len(image_names),
            formula_count=formula_count,
            warnings=warnings,
        )
    finally:
        reader.close()


def _main() -> None:
    import sys

    if len(sys.argv) < 3:
        print("用法: python -m backend.thesis_builder.docx_import <docx> <输出目录>")
        sys.exit(1)
    result = extract_docx_to_md(Path(sys.argv[1]), Path(sys.argv[2]))
    print(f"md: {result.md_path}")
    print(f"图片: {result.image_count}  公式: {result.formula_count}")
    for w in result.warnings:
        print(f"[!] {w}")


if __name__ == '__main__':
    _main()

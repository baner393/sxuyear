"""ThesisBuilder — 山西财经大学学年论文 Markdown→DOCX 自动排版引擎"""

__version__ = "0.1.0"

from .converter import convert
from .parser import parse_markdown, ParsedDocument
from .styles import ThesisConfig

__all__ = ["convert", "parse_markdown", "ParsedDocument", "ThesisConfig"]

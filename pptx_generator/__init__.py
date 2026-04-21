"""PPTX Generator — turn structured JSON, YAML, or Markdown into polished PPTX presentations."""

__version__ = "0.2.0"

from .generator import generate, BrandConfig, load_slides, Slide  # noqa: F401
from .markdown_parser import parse_markdown, parse_markdown_file  # noqa: F401

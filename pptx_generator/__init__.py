"""PPTX Generator — turn structured JSON, YAML, or Markdown into polished PPTX presentations."""

from importlib.metadata import version, PackageNotFoundError

try:
    __version__ = version("ai-pptx-generator")
except PackageNotFoundError:
    __version__ = "0.1.0"  # fallback for development mode

from .generator import generate, BrandConfig, load_slides, Slide  # noqa: F401
from .markdown_parser import parse_markdown, parse_markdown_file  # noqa: F401

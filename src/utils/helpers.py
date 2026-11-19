"""Shared utility functions for MCP Word server."""

import os
import sys


def normalize_path(path: str) -> str:
    """Normalize and expand a file path."""
    path = os.path.expanduser(path)
    path = os.path.expandvars(path)
    path = os.path.abspath(path)
    return path


def ensure_docx_extension(path: str) -> str:
    """Ensure the path ends with .docx."""
    if not path.lower().endswith(".docx"):
        path += ".docx"
    return path


def is_windows() -> bool:
    """Check if running on Windows."""
    return sys.platform == "win32"


def safe_int(value, default: int = 0) -> int:
    """Safely convert a value to int."""
    try:
        return int(value)
    except (TypeError, ValueError):
        return default


def pt_to_emu(pt: float) -> int:
    """Convert points to EMU (English Metric Units)."""
    return int(pt * 12700)


def inches_to_emu(inches: float) -> int:
    """Convert inches to EMU."""
    return int(inches * 914400)


def cm_to_emu(cm: float) -> int:
    """Convert centimeters to EMU."""
    return int(cm * 360000)

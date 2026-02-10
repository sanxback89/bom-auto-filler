"""
공통 텍스트 유틸리티 함수들
"""
import re
from typing import Optional


def clean_text(s: Optional[str]) -> str:
    """Remove unwanted line breaks/spaces while keeping original content as much as possible."""
    if not s:
        return ""
    s = str(s).replace("\r", " ").replace("\n", " ")
    s = re.sub(r"\s+", " ", s).strip()
    return s


def normalize_header(s: Optional[str]) -> str:
    s = clean_text(s).lower()
    s = re.sub(r"[^a-z0-9]+", "", s)  # keep alnum only
    return s


def clean_text_keep_newlines(s: Optional[str]) -> str:
    """
    Similar to clean_text(), but preserves newlines inside the string.
    Useful for header cells that are intentionally multi-line in the PDF.
    """
    if not s:
        return ""
    s = str(s).replace("\r", "\n")
    # normalize spaces around newlines
    s = re.sub(r"[ \t]+", " ", s)
    s = re.sub(r"\n[ \t]+", "\n", s)
    s = re.sub(r"[ \t]+\n", "\n", s)
    s = re.sub(r"\n{2,}", "\n", s)
    return s.strip()


def format_color_header_text(s: Optional[str]) -> str:
    """
    Format the color column header text we want to write into Excel (multi-line).
    We keep newlines if present and add a newline before a trailing CC number when possible.
    """
    t = clean_text_keep_newlines(s)
    if not t:
        return ""
    # If the header contains a long numeric token (BOM CC number), try to put it on its own line.
    m = re.search(r"(\b\d{9,}\b)", t)
    if m:
        cc = m.group(1)
        # IMPORTANT: keep trailing '-' if present in the PDF header (e.g., 'NY Athl Div -')
        before = t[: m.start(1)].rstrip()
        # Prefer splitting at " - " into its own line, but preserve the dash
        before = before.replace(" - ", " -\n")
        return (before + "\n" + cc).strip()
    return t.replace(" - ", " -\n").strip()

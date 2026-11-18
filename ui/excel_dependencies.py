"""Runtime preflight for Excel reporting workspaces.

This is intentionally just a check: installing packages at request time in a
Streamlit process is unreliable. Requirements are installed during deployment.
"""
from __future__ import annotations


def excel_dependency_error() -> str | None:
    """Return an actionable import failure, or None when Excel support loads."""
    try:
        from openpyxl import Workbook, load_workbook  # noqa: F401
    except ImportError as exc:
        return f"openpyxl could not be imported ({exc})."
    return None

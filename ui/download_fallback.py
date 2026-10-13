"""Browser-native fallback for small result downloads.

The normal st.download_button uses Streamlit's /media HTTP route. This option
embeds only SMALL outputs in a data URL, avoiding that route if its ASGI server
fails while serving files. Never inline big outputs into the page DOM.
"""

from __future__ import annotations

import base64
from html import escape

MAX_INLINE_DOWNLOAD_BYTES = 2 * 1024 * 1024  # 2 MiB (4/3 expansion in HTML)


def inline_download_link(data: bytes, filename: str, mime: str = "application/zip") -> str | None:
    """Return an HTML download anchor, or None for a large/empty payload."""
    if not isinstance(data, bytes) or not 0 < len(data) <= MAX_INLINE_DOWNLOAD_BYTES:
        return None
    if not filename or any(char in filename for char in ("/", "\\", "\r", "\n")):
        raise ValueError("Download filename must not contain paths or newlines")
    # Fixed, strictly scoped MIME: do not allow injection into an HTML attribute.
    if mime not in {"application/zip", "application/octet-stream", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"}:
        return None
    payload = base64.b64encode(data).decode("ascii")
    return (
        f'<a href="data:{mime};base64,{payload}" '
        f'download="{escape(filename, quote=True)}">'
        'Alternative: download directly from your browser (small file)</a>'
    )

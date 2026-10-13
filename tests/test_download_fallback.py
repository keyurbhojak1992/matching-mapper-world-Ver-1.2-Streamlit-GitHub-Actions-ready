from base64 import b64decode
from html.parser import HTMLParser
from ui.download_fallback import MAX_INLINE_DOWNLOAD_BYTES, inline_download_link


class LinkParser(HTMLParser):
    attrs = None
    def handle_starttag(self, tag, attrs):
        if tag == "a":
            self.attrs = dict(attrs)


def test_small_zip_link_downloads_exact_bytes_and_escaped_filename():
    data = b"PK\x03\x04binary-zip"
    text = inline_download_link(data, 'demand & "sales".zip')
    link = LinkParser()
    link.feed(text)
    assert link.attrs["download"] == 'demand & "sales".zip'
    prefix, encoded = link.attrs["href"].split(",", 1)
    assert prefix == "data:application/zip;base64"
    assert b64decode(encoded) == data


def test_large_and_empty_files_are_not_inlined():
    assert inline_download_link(b"", "sample.zip") is None
    assert inline_download_link(b"A" * (MAX_INLINE_DOWNLOAD_BYTES + 1), "sample.zip") is None


def test_rejects_unsafe_names_and_mime():
    import pytest
    with pytest.raises(ValueError):
        inline_download_link(b"123", "../hidden.zip")
    with pytest.raises(ValueError):
        inline_download_link(b"123", "foo\nbar.zip")
    assert inline_download_link(b"123", "file.zip", 'text/html" onclick="alert(1)') is None

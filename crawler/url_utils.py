from __future__ import annotations

from pathlib import PurePosixPath
from urllib.parse import parse_qsl, urlencode, urljoin, urlparse, urlunparse


SKIPPED_EXTENSIONS = {
    ".7z",
    ".avi",
    ".css",
    ".csv",
    ".doc",
    ".docx",
    ".gif",
    ".gz",
    ".ico",
    ".jpeg",
    ".jpg",
    ".js",
    ".json",
    ".mp3",
    ".mp4",
    ".pdf",
    ".png",
    ".ppt",
    ".pptx",
    ".rar",
    ".svg",
    ".tar",
    ".tsv",
    ".webm",
    ".webp",
    ".xls",
    ".xlsx",
    ".xml",
    ".zip",
}


def normalize_url(url: str, base_url: str | None = None) -> str | None:
    if any(character.isspace() for character in url):
        return None

    if base_url:
        url = urljoin(base_url, url)

    if any(character.isspace() for character in url):
        return None

    parsed = urlparse(url)
    if parsed.scheme not in {"http", "https"}:
        return None

    host = parsed.netloc.lower()
    if not host:
        return None

    path = parsed.path or "/"
    path = str(PurePosixPath(path))
    if parsed.path.endswith("/") and not path.endswith("/"):
        path += "/"

    query_pairs = [
        (key, value)
        for key, value in parse_qsl(parsed.query, keep_blank_values=True)
        if not key.lower().startswith("utm_")
    ]
    query = urlencode(sorted(query_pairs), doseq=True)

    return urlunparse((parsed.scheme.lower(), host, path, "", query, ""))


def is_allowed_domain(url: str, allowed_domains: list[str]) -> bool:
    return get_allowed_domain_key(url, allowed_domains) is not None


def get_allowed_domain_key(url: str, allowed_domains: list[str]) -> str | None:
    host = urlparse(url).netloc.lower()
    for domain in allowed_domains:
        if host == domain or host.endswith("." + domain):
            return domain
    return None


def looks_like_html_url(url: str) -> bool:
    path = urlparse(url).path.lower()
    return not any(path.endswith(ext) for ext in SKIPPED_EXTENSIONS)

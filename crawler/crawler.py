from __future__ import annotations

import argparse
import json
import time
from collections import deque
from dataclasses import dataclass
from html.parser import HTMLParser
from pathlib import Path
from typing import Iterable
from urllib import robotparser
from urllib.error import HTTPError, URLError
from urllib.parse import urlparse
from urllib.request import Request, urlopen

try:
    from url_utils import get_allowed_domain_key, is_allowed_domain, looks_like_html_url, normalize_url
except ImportError:
    from crawler.url_utils import get_allowed_domain_key, is_allowed_domain, looks_like_html_url, normalize_url


class LinkExtractor(HTMLParser):
    def __init__(self) -> None:
        super().__init__()
        self.links: list[str] = []

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        if tag.lower() != "a":
            return
        for key, value in attrs:
            if key.lower() == "href" and value:
                self.links.append(value)


@dataclass
class CrawlerConfig:
    seed_urls: list[str]
    allowed_domains: list[str]
    max_pages: int
    max_depth: int
    request_delay_seconds: float
    timeout_seconds: int
    user_agent: str
    output_dir: Path
    max_pages_per_domain: int | None = None
    excluded_url_contains: list[str] | None = None

    @classmethod
    def from_file(cls, path: Path) -> "CrawlerConfig":
        path = path.resolve()
        with path.open("r", encoding="utf-8") as file:
            data = json.load(file)
        base_dir = path.parent.parent
        output_dir = Path(data.get("output_dir", "data"))
        if not output_dir.is_absolute():
            output_dir = base_dir / output_dir
        return cls(
            seed_urls=data["seed_urls"],
            allowed_domains=[domain.lower() for domain in data["allowed_domains"]],
            max_pages=int(data.get("max_pages", 40)),
            max_depth=int(data.get("max_depth", 1)),
            request_delay_seconds=float(data.get("request_delay_seconds", 1.0)),
            timeout_seconds=int(data.get("timeout_seconds", 10)),
            user_agent=data.get("user_agent", "Student-Distributed-Search-Crawler/1.0"),
            output_dir=output_dir,
            max_pages_per_domain=(
                int(data["max_pages_per_domain"])
                if data.get("max_pages_per_domain") is not None
                else None
            ),
            excluded_url_contains=list(data.get("excluded_url_contains", [])),
        )


class ControlledCrawler:
    def __init__(self, config: CrawlerConfig) -> None:
        self.config = config
        self.raw_dir = config.output_dir / "raw"
        self.metadata_path = config.output_dir / "metadata.jsonl"
        self.visited: set[str] = set()
        self.robots_cache: dict[str, robotparser.RobotFileParser] = {}
        self.domain_page_counts: dict[str, int] = {}

    def crawl(self) -> int:
        self.raw_dir.mkdir(parents=True, exist_ok=True)
        self.metadata_path.parent.mkdir(parents=True, exist_ok=True)
        self.metadata_path.write_text("", encoding="utf-8")

        queue: deque[tuple[str, int]] = deque()
        for seed in self.config.seed_urls:
            normalized = normalize_url(seed)
            if normalized:
                queue.append((normalized, 0))

        downloaded = 0
        while queue and downloaded < self.config.max_pages:
            url, depth = queue.popleft()
            if not self._should_visit(url, depth):
                continue

            self.visited.add(url)
            page = self._fetch_html(url)
            if page is None:
                continue

            downloaded += 1
            domain = get_allowed_domain_key(url, self.config.allowed_domains) or urlparse(url).netloc.lower()
            self.domain_page_counts[domain] = self.domain_page_counts.get(domain, 0) + 1
            file_name = f"{downloaded:06d}.html"
            relative_file = f"raw/{file_name}"
            (self.raw_dir / file_name).write_bytes(page["body"])

            record = {
                "document_id": downloaded,
                "url": url,
                "file": relative_file,
                "source": urlparse(url).netloc,
                "crawl_depth": depth,
                "status": page["status"],
                "content_type": page["content_type"],
            }
            self._append_jsonl(self.metadata_path, record)
            print(f"[saved] {downloaded:03d} depth={depth} {url}")

            if depth < self.config.max_depth:
                for link in self._extract_links(page["text"], url):
                    if self._should_enqueue(link):
                        queue.append((link, depth + 1))

            time.sleep(self.config.request_delay_seconds)

        print(f"Crawl complete. Downloaded {downloaded} HTML pages.")
        return downloaded

    def _should_visit(self, url: str, depth: int) -> bool:
        return (
            url not in self.visited
            and depth <= self.config.max_depth
            and not self._is_excluded(url)
            and looks_like_html_url(url)
            and is_allowed_domain(url, self.config.allowed_domains)
            and self._below_domain_limit(url)
            and self._robots_allows(url)
        )

    def _should_enqueue(self, url: str) -> bool:
        return (
            url not in self.visited
            and not self._is_excluded(url)
            and looks_like_html_url(url)
            and is_allowed_domain(url, self.config.allowed_domains)
            and self._below_domain_limit(url)
        )

    def _is_excluded(self, url: str) -> bool:
        excluded_parts = self.config.excluded_url_contains or []
        url_lower = url.lower()
        return any(part.lower() in url_lower for part in excluded_parts)

    def _below_domain_limit(self, url: str) -> bool:
        if self.config.max_pages_per_domain is None:
            return True
        domain = get_allowed_domain_key(url, self.config.allowed_domains) or urlparse(url).netloc.lower()
        return self.domain_page_counts.get(domain, 0) < self.config.max_pages_per_domain

    def _fetch_html(self, url: str) -> dict[str, object] | None:
        request = Request(url, headers={"User-Agent": self.config.user_agent})
        try:
            with urlopen(request, timeout=self.config.timeout_seconds) as response:
                content_type = response.headers.get("Content-Type", "")
                if "text/html" not in content_type.lower():
                    print(f"[skip] non-html {url} ({content_type})")
                    return None
                body = response.read()
                charset = response.headers.get_content_charset() or "utf-8"
                text = body.decode(charset, errors="replace")
                return {
                    "body": body,
                    "text": text,
                    "status": getattr(response, "status", 200),
                    "content_type": content_type,
                }
        except Exception as exc:
            print(f"[error] {url} -> {exc}")
            return None

    def _extract_links(self, html: str, base_url: str) -> Iterable[str]:
        parser = LinkExtractor()
        try:
            parser.feed(html)
        except Exception:
            return []

        links: list[str] = []
        for href in parser.links:
            normalized = normalize_url(href, base_url)
            if normalized:
                links.append(normalized)
        return links

    def _robots_allows(self, url: str) -> bool:
        parsed = urlparse(url)
        root = f"{parsed.scheme}://{parsed.netloc}"
        if root not in self.robots_cache:
            rp = robotparser.RobotFileParser()
            rp.set_url(root + "/robots.txt")
            try:
                rp.read()
            except Exception:
                pass
            self.robots_cache[root] = rp
        return self.robots_cache[root].can_fetch(self.config.user_agent, url)

    @staticmethod
    def _append_jsonl(path: Path, record: dict[str, object]) -> None:
        with path.open("a", encoding="utf-8") as file:
            file.write(json.dumps(record, ensure_ascii=True) + "\n")


def main() -> None:
    parser = argparse.ArgumentParser(description="Controlled academic web crawler")
    parser.add_argument("--config", required=True, help="Path to crawler_config.json")
    args = parser.parse_args()

    config = CrawlerConfig.from_file(Path(args.config))
    ControlledCrawler(config).crawl()


if __name__ == "__main__":
    main()

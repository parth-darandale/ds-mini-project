from __future__ import annotations

import argparse
import json
from html.parser import HTMLParser
from pathlib import Path


class VisibleTextParser(HTMLParser):
    SKIP_TAGS = {"script", "style", "noscript", "svg", "canvas"}

    def __init__(self) -> None:
        super().__init__()
        self.skip_depth = 0
        self.in_title = False
        self.title_parts: list[str] = []
        self.text_parts: list[str] = []

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        tag = tag.lower()
        if tag in self.SKIP_TAGS:
            self.skip_depth += 1
        if tag == "title":
            self.in_title = True
        if tag in {"p", "div", "br", "li", "h1", "h2", "h3"}:
            self.text_parts.append(" ")

    def handle_endtag(self, tag: str) -> None:
        tag = tag.lower()
        if tag in self.SKIP_TAGS and self.skip_depth > 0:
            self.skip_depth -= 1
        if tag == "title":
            self.in_title = False
        if tag in {"p", "div", "li", "h1", "h2", "h3"}:
            self.text_parts.append(" ")

    def handle_data(self, data: str) -> None:
        if self.skip_depth > 0:
            return
        text = " ".join(data.split())
        if not text:
            return
        if self.in_title:
            self.title_parts.append(text)
        self.text_parts.append(text)

    @property
    def title(self) -> str:
        return " ".join(self.title_parts).strip()

    @property
    def visible_text(self) -> str:
        return " ".join(" ".join(self.text_parts).split())


def read_jsonl(path: Path) -> list[dict[str, object]]:
    records: list[dict[str, object]] = []
    if not path.exists():
        return records
    with path.open("r", encoding="utf-8") as file:
        for line in file:
            line = line.strip()
            if line:
                records.append(json.loads(line))
    return records


def parse_documents(data_dir: Path) -> int:
    metadata_path = data_dir / "metadata.jsonl"
    output_path = data_dir / "documents.jsonl"
    metadata = read_jsonl(metadata_path)
    output_path.write_text("", encoding="utf-8")

    written = 0
    with output_path.open("a", encoding="utf-8") as output:
        for record in metadata:
            html_path = data_dir / str(record["file"])
            if not html_path.exists():
                print(f"[skip] missing file {html_path}")
                continue
            html = html_path.read_text(encoding="utf-8", errors="replace")
            parser = VisibleTextParser()
            parser.feed(html)
            content = parser.visible_text
            if not content:
                print(f"[skip] no visible text in {html_path}")
                continue

            document = {
                "document_id": int(record["document_id"]),
                "title": parser.title or str(record["url"]),
                "url": str(record["url"]),
                "source": str(record.get("source", "")),
                "crawl_depth": int(record.get("crawl_depth", 0)),
                "content": content,
            }
            output.write(json.dumps(document, ensure_ascii=True) + "\n")
            written += 1

    print(f"Parsed {written} logical documents into {output_path}")
    return written


def main() -> None:
    parser = argparse.ArgumentParser(description="Parse raw HTML into documents.jsonl")
    parser.add_argument("--data-dir", default="data", help="Data directory containing raw/ and metadata.jsonl")
    args = parser.parse_args()
    parse_documents(Path(args.data_dir))


if __name__ == "__main__":
    main()


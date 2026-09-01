from __future__ import annotations

import argparse
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT))

from crawler.crawler import ControlledCrawler, CrawlerConfig
from indexing.index_builder import build_indexes
from indexing.partitioner import partition_documents
from parser.html_parser import parse_documents


def main() -> None:
    parser = argparse.ArgumentParser(description="Run Member 1 offline search-engine pipeline")
    parser.add_argument("--config", default=str(PROJECT_ROOT / "config" / "crawler_config.json"))
    parser.add_argument("--num-shards", type=int, default=3)
    parser.add_argument("--skip-crawl", action="store_true", help="Use existing data/raw and metadata.jsonl")
    args = parser.parse_args()

    config = CrawlerConfig.from_file(Path(args.config))
    data_dir = config.output_dir

    if not args.skip_crawl:
        ControlledCrawler(config).crawl()

    parse_documents(data_dir)
    partition_documents(data_dir / "documents.jsonl", data_dir / "shards", args.num_shards)
    build_indexes(data_dir / "shards", PROJECT_ROOT / "indexes")
    print("Member 1 pipeline complete.")


if __name__ == "__main__":
    main()

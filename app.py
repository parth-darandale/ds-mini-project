from __future__ import annotations

import argparse
import json
import math
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(PROJECT_ROOT))

from crawler.crawler import ControlledCrawler, CrawlerConfig
from indexing.index_builder import build_indexes
from indexing.partitioner import partition_documents
from indexing.text_utils import tokenize
from parser.html_parser import parse_documents


def run_pipeline(args: argparse.Namespace) -> None:
    config = CrawlerConfig.from_file(Path(args.config))
    data_dir = config.output_dir

    if not args.skip_crawl:
        ControlledCrawler(config).crawl()

    parse_documents(data_dir)
    partition_documents(data_dir / "documents.jsonl", data_dir / "shards", args.num_shards)
    build_indexes(data_dir / "shards", PROJECT_ROOT / "indexes")
    print("Pipeline complete.")


def show_stats(args: argparse.Namespace) -> None:
    data_dir = PROJECT_ROOT / "data"
    indexes_dir = PROJECT_ROOT / "indexes"
    metadata_count = count_jsonl(data_dir / "metadata.jsonl")
    document_count = count_jsonl(data_dir / "documents.jsonl")
    raw_count = len(list((data_dir / "raw").glob("*.html")))

    print("Distributed Search Engine - Offline Corpus Stats")
    print(f"Raw HTML pages: {raw_count}")
    print(f"Metadata records: {metadata_count}")
    print(f"Parsed documents: {document_count}")

    stats_path = indexes_dir / "global_stats.json"
    if stats_path.exists():
        stats = json.loads(stats_path.read_text(encoding="utf-8"))
        print(f"Global indexed documents: {stats['total_documents']}")
        for node in stats["nodes"]:
            print(
                f"{node['node_id']}: "
                f"{node['document_count']} documents, {node['unique_terms']} unique terms"
            )


def search_local(args: argparse.Namespace) -> None:
    query_tokens = tokenize(args.query)
    if not query_tokens:
        print("Please enter a query with searchable terms.")
        return

    stats = json.loads((PROJECT_ROOT / "indexes" / "global_stats.json").read_text(encoding="utf-8"))
    global_idf = stats["global_idf"]
    query_weights = build_query_weights(query_tokens, global_idf)
    query_norm = vector_norm(query_weights.values())
    results: list[dict[str, object]] = []

    for node_dir in sorted((PROJECT_ROOT / "indexes").glob("node_*")):
        index_path = node_dir / "index.json"
        if not index_path.exists():
            continue
        node_index = json.loads(index_path.read_text(encoding="utf-8"))
        results.extend(search_node_index(node_dir.name, node_index, query_tokens, query_weights, query_norm))

    results.sort(key=lambda item: float(item["score"]), reverse=True)
    for rank, result in enumerate(results[: args.top_k], start=1):
        print(f"\n{rank}. {result['title']}")
        print(result["url"])
        print(f"Score: {result['score']:.4f} | Node: {result['node_id']}")
        print(result["snippet"])


def search_node_index(
    node_id: str,
    node_index: dict[str, object],
    query_tokens: list[str],
    query_weights: dict[str, float],
    query_norm: float,
) -> list[dict[str, object]]:
    inverted_index = node_index["inverted_index"]
    document_lengths = node_index["document_lengths"]
    documents = node_index["documents"]
    candidates: dict[str, dict[str, float]] = {}

    for term in query_tokens:
        for posting in inverted_index.get(term, []):
            doc_id = str(posting["document_id"])
            doc_length = max(1, int(document_lengths.get(doc_id, 1)))
            tf = int(posting["frequency"]) / doc_length
            candidates.setdefault(doc_id, {})[term] = tf * query_weights.get(term, 0.0)

    results: list[dict[str, object]] = []
    for doc_id, doc_weights in candidates.items():
        doc_norm = vector_norm(doc_weights.values())
        if query_norm == 0 or doc_norm == 0:
            continue
        dot_product = sum(query_weights.get(term, 0.0) * weight for term, weight in doc_weights.items())
        score = dot_product / (query_norm * doc_norm)
        document = documents[doc_id]
        results.append(
            {
                "document_id": int(doc_id),
                "title": document["title"],
                "url": document["url"],
                "node_id": node_id,
                "score": score,
                "snippet": make_snippet(str(document["content"]), query_tokens),
            }
        )

    return results


def build_query_weights(tokens: list[str], global_idf: dict[str, float]) -> dict[str, float]:
    counts: dict[str, int] = {}
    for token in tokens:
        counts[token] = counts.get(token, 0) + 1
    total = max(1, len(tokens))
    return {
        token: (count / total) * float(global_idf.get(token, 0.0))
        for token, count in counts.items()
    }


def make_snippet(content: str, tokens: list[str], size: int = 220) -> str:
    content_lower = content.lower()
    first_match = min(
        [content_lower.find(token) for token in tokens if content_lower.find(token) >= 0],
        default=0,
    )
    start = max(0, first_match - size // 3)
    snippet = content[start : start + size].replace("\n", " ")
    return " ".join(snippet.split())


def vector_norm(values: object) -> float:
    return math.sqrt(sum(float(value) * float(value) for value in values))


def count_jsonl(path: Path) -> int:
    if not path.exists():
        return 0
    return sum(1 for line in path.open("r", encoding="utf-8") if line.strip())


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Distributed Search Engine offline app")
    subparsers = parser.add_subparsers(dest="command", required=True)

    pipeline = subparsers.add_parser("pipeline", help="Crawl, parse, shard, and index documents")
    pipeline.add_argument("--config", default=str(PROJECT_ROOT / "config" / "crawler_config_large.json"))
    pipeline.add_argument("--num-shards", type=int, default=3)
    pipeline.add_argument("--skip-crawl", action="store_true")
    pipeline.set_defaults(func=run_pipeline)

    stats = subparsers.add_parser("stats", help="Show corpus and index statistics")
    stats.set_defaults(func=show_stats)

    search = subparsers.add_parser("search", help="Run a local search across generated indexes")
    search.add_argument("query")
    search.add_argument("--top-k", type=int, default=10)
    search.set_defaults(func=search_local)

    return parser


def main() -> None:
    parser = build_parser()
    args = parser.parse_args()
    args.func(args)


if __name__ == "__main__":
    main()


from __future__ import annotations

import argparse
import json
import math
from collections import Counter, defaultdict
from pathlib import Path

try:
    from text_utils import tokenize
except ImportError:
    from indexing.text_utils import tokenize


def read_jsonl(path: Path) -> list[dict[str, object]]:
    records: list[dict[str, object]] = []
    with path.open("r", encoding="utf-8") as file:
        for line in file:
            line = line.strip()
            if line:
                records.append(json.loads(line))
    return records


def build_index_for_documents(documents: list[dict[str, object]]) -> dict[str, object]:
    inverted_index: dict[str, list[dict[str, object]]] = defaultdict(list)
    document_store: dict[str, dict[str, object]] = {}
    document_lengths: dict[str, int] = {}
    local_df: Counter[str] = Counter()

    for document in documents:
        document_id = int(document["document_id"])
        doc_key = str(document_id)
        tokens = tokenize(str(document.get("content", "")))
        positions_by_term: dict[str, list[int]] = defaultdict(list)
        for position, token in enumerate(tokens):
            positions_by_term[token].append(position)

        document_lengths[doc_key] = len(tokens)
        document_store[doc_key] = {
            "document_id": document_id,
            "title": document.get("title", ""),
            "url": document.get("url", ""),
            "source": document.get("source", ""),
            "content": document.get("content", ""),
        }

        for term, positions in positions_by_term.items():
            local_df[term] += 1
            inverted_index[term].append(
                {
                    "document_id": document_id,
                    "frequency": len(positions),
                    "positions": positions,
                }
            )

    return {
        "document_count": len(documents),
        "document_lengths": document_lengths,
        "documents": document_store,
        "local_df": dict(local_df),
        "inverted_index": dict(sorted(inverted_index.items())),
    }


def build_indexes(shards_dir: Path, index_dir: Path) -> None:
    shard_paths = sorted(shards_dir.glob("shard_*.jsonl"))
    if not shard_paths:
        raise FileNotFoundError(f"No shard_*.jsonl files found in {shards_dir}")

    index_dir.mkdir(parents=True, exist_ok=True)
    global_df: Counter[str] = Counter()
    total_documents = 0
    node_summaries: list[dict[str, object]] = []

    for shard_number, shard_path in enumerate(shard_paths, start=1):
        documents = read_jsonl(shard_path)
        index = build_index_for_documents(documents)
        node_dir = index_dir / f"node_{shard_number}"
        node_dir.mkdir(parents=True, exist_ok=True)
        index_path = node_dir / "index.json"
        index_path.write_text(json.dumps(index, indent=2, ensure_ascii=True), encoding="utf-8")

        total_documents += int(index["document_count"])
        global_df.update(index["local_df"])
        node_summaries.append(
            {
                "node_id": f"node_{shard_number}",
                "shard_file": str(shard_path).replace("\\", "/"),
                "index_file": str(index_path).replace("\\", "/"),
                "document_count": index["document_count"],
                "unique_terms": len(index["inverted_index"]),
            }
        )
        print(f"[index] {index_path} ({index['document_count']} docs)")

    global_idf = {
        term: math.log((1 + total_documents) / (1 + df)) + 1
        for term, df in sorted(global_df.items())
    }
    global_stats = {
        "total_documents": total_documents,
        "global_df": dict(sorted(global_df.items())),
        "global_idf": global_idf,
        "nodes": node_summaries,
    }
    stats_path = index_dir / "global_stats.json"
    stats_path.write_text(json.dumps(global_stats, indent=2, ensure_ascii=True), encoding="utf-8")
    print(f"[stats] {stats_path} ({total_documents} total docs)")


def main() -> None:
    parser = argparse.ArgumentParser(description="Build local inverted indexes and global stats")
    parser.add_argument("--shards-dir", default="data/shards")
    parser.add_argument("--index-dir", default="indexes")
    args = parser.parse_args()
    build_indexes(Path(args.shards_dir), Path(args.index_dir))


if __name__ == "__main__":
    main()


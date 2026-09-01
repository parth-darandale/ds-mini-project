from __future__ import annotations

import argparse
import json
from pathlib import Path


def read_documents(path: Path) -> list[dict[str, object]]:
    documents: list[dict[str, object]] = []
    with path.open("r", encoding="utf-8") as file:
        for line in file:
            line = line.strip()
            if line:
                documents.append(json.loads(line))
    return documents


def partition_documents(documents_path: Path, output_dir: Path, num_shards: int) -> list[Path]:
    documents = read_documents(documents_path)
    output_dir.mkdir(parents=True, exist_ok=True)

    shard_paths = [output_dir / f"shard_{index}.jsonl" for index in range(1, num_shards + 1)]
    for path in shard_paths:
        path.write_text("", encoding="utf-8")

    shard_files = [path.open("a", encoding="utf-8") for path in shard_paths]
    try:
        total_documents = len(documents)
        shard_size = max(1, (total_documents + num_shards - 1) // num_shards)
        for position, document in enumerate(documents):
            shard_index = min(position // shard_size, num_shards - 1)
            shard_files[shard_index].write(json.dumps(document, ensure_ascii=True) + "\n")
    finally:
        for file in shard_files:
            file.close()

    print(f"Partitioned {len(documents)} documents into {num_shards} shards.")
    for path in shard_paths:
        print(f"[shard] {path}")
    return shard_paths


def main() -> None:
    parser = argparse.ArgumentParser(description="Split documents into search-node shards")
    parser.add_argument("--documents", default="data/documents.jsonl")
    parser.add_argument("--output-dir", default="data/shards")
    parser.add_argument("--num-shards", type=int, default=3)
    args = parser.parse_args()
    partition_documents(Path(args.documents), Path(args.output_dir), args.num_shards)


if __name__ == "__main__":
    main()

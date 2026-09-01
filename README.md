# Fault-Tolerant Distributed Search Engine - Offline Pipeline

This is the Member 1 implementation for FA-1 of the distributed search engine
mini project. It prepares the offline corpus, partitions documents, and builds
the indexes that the distributed search-node layer will use.

Member 1 is responsible for:

1. Controlled crawling of public technology pages.
2. Saving raw HTML and metadata.
3. Parsing raw HTML into logical searchable documents.
4. Splitting documents into shards for multiple search nodes.
5. Building one local inverted index per shard.
6. Computing global TF-IDF statistics used by all search nodes.

## Why This Is Distributed

The system uses document-based partitioning. The full document collection is not
kept as one central index. Instead:

```text
documents.jsonl
      |
      v
shard_1.jsonl -> indexes/node_1/index.json
shard_2.jsonl -> indexes/node_2/index.json
shard_3.jsonl -> indexes/node_3/index.json
```

Each search node in Member 2's part will load only its own shard and local
index. The coordinator can then query the search nodes in parallel using gRPC.

## Folder Structure

```text
distributed-search-engine/
├── config/
│   └── crawler_config.json
├── crawler/
│   ├── crawler.py
│   └── url_utils.py
├── parser/
│   └── html_parser.py
├── indexing/
│   ├── index_builder.py
│   ├── partitioner.py
│   └── text_utils.py
├── scripts/
│   └── run_pipeline.py
├── data/
│   ├── raw/
│   ├── metadata.jsonl
│   ├── documents.jsonl
│   └── shards/
└── indexes/
```

## Run The Full Pipeline

From this folder:

```powershell
python app.py pipeline --config config/crawler_config.json
```

If `python` is not configured, try:

```powershell
py -3 app.py pipeline --config config/crawler_config.json
```

For a quick test run:

```powershell
python app.py pipeline --config config/crawler_config_test.json
```

For the 1,000-document FA-1 corpus run:

```powershell
python app.py pipeline --config config/crawler_config_large.json
```

The large config is capped at 1,000 pages total and 100 pages per website. It
may take around 30 minutes because the crawler uses a delay between requests and
checks `robots.txt`.

## Run Steps Separately

Crawl pages:

```powershell
python crawler/crawler.py --config config/crawler_config.json
```

Parse HTML:

```powershell
python parser/html_parser.py --data-dir data
```

Partition documents:

```powershell
python indexing/partitioner.py --documents data/documents.jsonl --output-dir data/shards --num-shards 3
```

Build indexes:

```powershell
python indexing/index_builder.py --shards-dir data/shards --index-dir indexes
```

## App Commands

Show corpus/index statistics:

```powershell
python app.py stats
```

Try a local search over the generated indexes:

```powershell
python app.py search "distributed systems"
python app.py search "parallel query processing" --top-k 5
```

The local search command is only a simple verification tool for Member 1. Member
2 will implement the real distributed query flow using gRPC search nodes and a
coordinator.

## Outputs For Member 2

Member 2 should use these files:

```text
data/shards/shard_1.jsonl
data/shards/shard_2.jsonl
data/shards/shard_3.jsonl
indexes/node_1/index.json
indexes/node_2/index.json
indexes/node_3/index.json
indexes/global_stats.json
```

The partitioning is range-based for FA-1. With three shards, the first range of
documents goes to `shard_1`, the next range goes to `shard_2`, and the final
range goes to `shard_3`.

## Suggested Demo

For development, use 30-50 pages. For FA-1 final demo, increase `max_pages` in
`config/crawler_config.json` to around 200 if the crawl is stable.

For a later 10,000-document corpus, increase `max_pages` to `10000` and
`max_pages_per_domain` to `1000`. If a site blocks crawling through `robots.txt`
or returns many errors, the crawler will skip those pages and continue with the
remaining allowed domains.

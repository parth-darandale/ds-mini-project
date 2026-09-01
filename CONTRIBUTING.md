# Contributing

This is an academic Distributed Systems mini project. Contributions are welcome
through forks, pull requests, or collaborator access from the repository owner.

## Development Flow

1. Keep crawler configuration changes in `config/`.
2. Keep raw crawled pages in `data/raw/`.
3. Regenerate derived files with:

```powershell
python app.py pipeline --config config/crawler_config_large.json
```

4. Check the corpus and index counts with:

```powershell
python app.py stats
```

5. Try a local search before handing data to the distributed search-node layer:

```powershell
python app.py search "parallel query processing"
```

## Notes

The repository is licensed under MIT so others can view, use, and modify the
code. Direct push access on GitHub still requires the repository owner to add
collaborators.


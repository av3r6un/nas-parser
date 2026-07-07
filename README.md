# NAS Parser

Production ETL для прайс-листов NAS Crystal.

Проект читает Excel-прайсы, преобразует их в единый каталог для импорта и
дополняет результат отдельным аудитом качества готового `catalog.xlsx`.

## Документация

- [Architecture](docs/architecture.md)
- [Development History](docs/development-history.md)
- [Parser Specification](docs/parser-spec.md)
- [Business Rules](docs/business-rules.md)
- [Audit](docs/audit.md)
- [CLI](docs/cli.md)
- [Testing](docs/testing.md)
- [Project Structure](docs/project-structure.md)
- [Roadmap](docs/roadmap.md)
- [Contributing](CONTRIBUTING.md)
- [Changelog](CHANGELOG.md)

## Quick Start

```bash
uv sync
uv run python -m nas_parser.main
uv run python -m nas_parser.audit output/catalog.xlsx
```

If you already have dependencies installed in a local Python environment, you can run
the same entry points directly with `python`:

```bash
python -m nas_parser.main
python -m nas_parser.audit output/catalog.xlsx
```

## Requirements

- Python `3.14+`
- `uv` as the primary workflow tool

## Result

- export file: `output/catalog.xlsx`
- logs: `logs/info.log`, `logs/warnings.log`, `logs/errors.log`

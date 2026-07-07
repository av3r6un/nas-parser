# Структура проекта

## Верхний уровень

```text
docs/                 документация проекта
input/                входные Excel-файлы
logs/                 runtime-логи
nas_parser/           основной пакет
output/               итоговый catalog.xlsx
reference/            reference-файлы
tests/                unit и integration tests
CHANGELOG.md          история релизных изменений
CONTRIBUTING.md       правила развития проекта
README.md             витрина проекта
main.py               верхнеуровневый entrypoint
pyproject.toml        метаданные проекта
requirements.txt      резервный путь установки
```

## Пакет `nas_parser`

```text
nas_parser/
├── __init__.py
├── audit.py
├── business.py
├── config.py
├── domain.py
├── export.py
├── main.py
├── pipeline.py
├── report.py
├── source.py
├── validation.py
├── parsers/
│   ├── __init__.py
│   ├── base.py
│   ├── cut.py
│   ├── k9.py
│   └── registry.py
├── readers/
│   ├── __init__.py
│   ├── base.py
│   └── excel.py
└── references/
    ├── __init__.py
    ├── base.py
    └── colors.py
```

## Где искать что

### Входной Excel

- `nas_parser/readers/`
- `nas_parser/source.py`

### Разбор форматов

- `nas_parser/parsers/cut.py`
- `nas_parser/parsers/k9.py`
- `nas_parser/parsers/registry.py`

### Бизнес-обогащение

- `nas_parser/business.py`

### Валидация

- `nas_parser/validation.py`

### Экспорт

- `nas_parser/export.py`

### Координация ETL

- `nas_parser/pipeline.py`
- `nas_parser/main.py`

### Аудит результата

- `nas_parser/audit.py`

### Тесты

- `tests/test_cut_parser.py`
- `tests/test_k9_parser.py`
- `tests/test_business.py`
- `tests/test_validation.py`
- `tests/test_export.py`
- `tests/test_pipeline.py`
- `tests/test_audit.py`

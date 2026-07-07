# CLI

## Подготовка окружения

```bash
uv sync
```

If dependencies are already installed in a regular Python environment, `uv` is optional.
You can run the same entry points directly with `python`.

## Основной ETL запуск

```bash
uv run python -m nas_parser.main
python -m nas_parser.main
python main.py
```

Результат:

- создается `output/catalog.xlsx`;
- формируются логи в `logs/`.

## Аудит готового каталога

```bash
uv run python -m nas_parser.audit output/catalog.xlsx
python -m nas_parser.audit output/catalog.xlsx
```

Команда выводит:

- краткую статистику;
- проблемные строки каталога, если они найдены.

## Запуск тестов

```bash
uv run python -m unittest discover -s tests
python -m unittest discover -s tests
```

## Альтернативный запуск

Если нужен запуск через файл верхнего уровня:

```bash
uv run python main.py
python main.py
```

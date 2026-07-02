# NAS Crystal Parser

Парсер Excel-прайсов NAS Crystal. Проект читает исходные таблицы из `input/`,
обогащает товары кодами цветов из справочника `reference/`, валидирует записи и
формирует итоговый файл для импорта в `output/catalog.xlsx`.

## Что поддерживается

- 12cut hot/non-hotfix файлы.
- 16cut hot/non-hotfix файлы.
- K9 / пришивные стразы.
- Подстановка кода цвета из справочника, если код найден.
- Построчные логи `info`, `warnings`, `errors`.

## Структура папок

```text
input/       исходные Excel-файлы прайсов
reference/   справочники, например colorcode-articul.xlsx
output/      итоговый catalog.xlsx
logs/        info.log, warnings.log, errors.log
nas_parser/  код приложения
tests/       тесты
```

Папки `input/`, `reference/`, `output/`, `logs/` хранятся в репозитории через
`.gitkeep`. Сами рабочие `.xlsx` и `.log` файлы не коммитятся.

## Установка

Нужен Python `3.14+`. Основной способ запуска использует `uv`.

```bash
uv sync
```

Если `uv` не установлен, можно поставить зависимости через `pip`:

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Подготовка данных

1. Положите прайсы в папку `input/`.
2. Для 12cut и 16cut имя файла должно содержать `12cut` или `16cut`.
3. Для K9 имя файла должно содержать `K9` или `k9`.
4. Положите справочник цветов в `reference/colorcode-articul.xlsx`.

Справочник цветов должен содержать колонки:

```text
Цвет
Артикул цвета
```

Если код цвета найден, в финальный файл попадет код, например `016`. Если кода
нет, останется буквенное название цвета.

## Запуск

```bash
uv run python main.py
```

После запуска в консоли появится краткая сводка:

```text
files_found=11 files_processed=11 files_skipped=0 records_created=1914 records_12cut=912 records_16cut=135 records_k9=867 info=13 warnings=98 errors=0 output_file=output/catalog.xlsx
```

## Результат

Итоговый файл создается здесь:

```text
output/catalog.xlsx
```

В финальный файл передаются кодовые значения:

- `Фиксация`: `Hot`, `Non`, `K9`.
- `Грани`: `12cut`, `16cut`.
- `Цвет`: код из справочника, если найден; иначе исходное название цвета.

## Логи

Подробности запуска пишутся в папку `logs/`:

```text
logs/info.log
logs/warnings.log
logs/errors.log
```

Каждая запись пишется отдельной строкой. Например:

```text
Missing quantity for K9.xlsx:пришивные:783
```

## Тесты

```bash
uv run python -m unittest discover -s tests
```

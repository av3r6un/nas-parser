# Testing

## Основная команда

```bash
uv run python -m unittest discover -s tests
python -m unittest discover -s tests
```

Проект использует `unittest`.

## Что покрыто

### Domain

- базовая модель `ProductRecord`.

### Readers

- базовый reader;
- чтение Excel.

### Parsers

- `ParserBase`;
- `ParserRegistry`;
- `CutParser`;
- `K9Parser`.

### Business

- построение `sku`;
- построение `name`;
- построение `category`;
- применение `color_code`.

### Validation

- проверки заполненности и корректности записей.

### Export

- маппинг `ProductRecord` в строки итогового Excel.

### Pipeline

- сквозной запуск от input до `catalog.xlsx`.

### Report

- форматирование `RunReport`;
- запись логов.

### Audit

- статистика качества готового каталога;
- CLI аудита;
- детализация проблемных строк.

## Принципы тестирования

- тесты закрепляют уже согласованное поведение;
- production-polish изменения должны быть локальными;
- новые тесты добавляются рядом с тем слоем, который реально меняется;
- рефакторинг ради тестов не является целью сам по себе.

# Changelog

## v1.0.0

Первый production-ready релиз NAS Parser.

### Core ETL

- реализован полный ETL pipeline от входных Excel до `output/catalog.xlsx`;
- добавлены `ExcelReader`, `CutParser`, `K9Parser`, `ProductEnricher`,
  `ProductValidator`, `ExcelExporter`, `Pipeline`;
- закреплена центральная модель `ProductRecord`.

### Format support

- поддержаны `12cut` и `16cut` hot/non-hotfix прайсы;
- поддержаны `K9` и `K9_premium` файлы;
- добавлена поддержка готового `Артикул` для cut-файлов;
- добавлена поддержка `mix` для перестановки tail columns;
- добавлена поддержка Excel-формул в price и quantity.

### Audit and tooling

- добавлен отдельный `CatalogAudit` для готового каталога;
- добавлен CLI для аудита;
- добавлена детализация проблемных строк в audit output;
- устранен `RuntimeWarning` при `uv run python -m nas_parser.main`.

### Documentation

- добавлен каталог `docs/` с архитектурной и эксплуатационной документацией;
- `README.md` сокращен до роли витрины проекта;
- добавлен `CONTRIBUTING.md`.

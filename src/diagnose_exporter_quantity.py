from decimal import Decimal
from pathlib import Path
import sys

# Добавляем корневую директорию проекта в sys.path, чтобы можно было импортировать модули проекта
# Это временное изменение sys.path только для этого скрипта и не влияет на проект.
project_root = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(project_root))

from nas_parser.domain import ProductRecord
from nas_parser.export import ExcelExporter
from nas_parser.report import RunReport

# Имитация ProductRecord, поступающего из K9Parser
def create_dummy_product_record(quantity_value, file_name="dummy_k9.xlsx"):
    return ProductRecord(
        price=Decimal('10.00'),
        quantity=quantity_value,
        color="Crystal",
        size="SS16",
        shape="Round",
        fixation="sew",
        cut=None,
        source_file=Path(file_name),
        source_sheet="Sheet1",
        source_row=1,
        parser_name="k9",
        sku=None, name=None, category=None, color_code=None
    )

def diagnose_exporter_quantity():
    print("--- Анализ ProductRecord.quantity после K9Parser (имитация) ---")

    # Сценарий 1: Количество - число
    record_num = create_dummy_product_record(Decimal('72'))
    print(f"\n1. Имитация ProductRecord с quantity=Decimal('72'):")
    print(f"   type(record.quantity): {type(record_num.quantity)}")
    print(f"   repr(record.quantity): {repr(record_num.quantity)}")
    print(f"   isinstance(record.quantity, Decimal): {isinstance(record_num.quantity, Decimal)}")
    print(f"   Значение quantity: {record_num.quantity}")

    # Сценарий 2: Количество - строка-формула
    record_formula_str = create_dummy_product_record("=G17-C17")
    print(f"\n2. Имитация ProductRecord с quantity='=G17-C17' (строка-формула):")
    print(f"   type(record.quantity): {type(record_formula_str.quantity)}")
    print(f"   repr(record.quantity): {repr(record_formula_str.quantity)}")
    print(f"   isinstance(record.quantity, Decimal): {isinstance(record_formula_str.quantity, Decimal)}")
    print(f"   Значение quantity: {record_formula_str.quantity}")

    # Сценарий 3: Количество - строка (не формула)
    record_text_str = create_dummy_product_record("Нет в наличии")
    print(f"\n3. Имитация ProductRecord с quantity='Нет в наличии' (строка-текст):")
    print(f"   type(record.quantity): {type(record_text_str.quantity)}")
    print(f"   repr(record.quantity): {repr(record_text_str.quantity)}")
    print(f"   isinstance(record.quantity, Decimal): {isinstance(record_text_str.quantity, Decimal)}")
    print(f"   Значение quantity: {record_text_str.quantity}")

    print("\n--- Анализ того, что ExcelExporter записывает в cell.value ---")

    # Создаем фиктивный ExcelExporter для демонстрации метода _record_to_row
    # Output file не будет создан, мы только вызываем внутренний метод.
    exporter = ExcelExporter(Path("temp_output.xlsx"))
    
    # Чтобы не вызывать реальный метод export, который создает файл,
    # мы имитируем вызов _record_to_row и показываем, что он возвращает.
    
    # Для quantity, ExcelExporter.py (на основе последнего анализа)
    # вызывает _decimal_value для парсеров Cut/K9.
    # Но в _record_to_row он использует record.quantity как есть.
    
    # Имитация _record_to_row для разных типов quantity
    # (для точного поведения нужно знать, какой индекс у quantity в HEADER)
    # По умолчанию Quantity - 4-ый элемент (индекс 3) в моем последнем HEADER.
    # Предположим, что quantity находится в 4-й колонке (индекс 3) для демонстрации.
    
    # Найдем фактический индекс quantity в ExcelExporter.HEADER
    try:
        from nas_parser.export import HEADER
        quantity_header_name = "Количество" # Или "К-во" или как в шаблоне 1С
        if quantity_header_name in HEADER:
            quantity_column_pos = HEADER.index(quantity_header_name)
        else:
            # Fallback if "Количество" is not directly in HEADER, use a guess or default
            # based on prior discussions, quantity is typically around 6th-8th column in output
            # For now, let's use a dummy index for demonstration if not found.
            quantity_column_pos = 5 # arbitrary index for demonstration
            print(f"  Внимание: '{quantity_header_name}' не найден в HEADER ExcelExporter. Используется фиктивный индекс {quantity_column_pos}.")
    except ImportError:
        print("  Не удалось импортировать HEADER из nas_parser.export. Используется фиктивный индекс 5.")
        quantity_column_pos = 5


    print(f"\nПоведение _record_to_row (колонка Quantity - индекс {quantity_column_pos}):")

    # Имитация результата _record_to_row для record_num
    row_values_num = exporter._record_to_row(record_num)
    print(f"  Для Decimal Quantity ({record_num.quantity}):")
    print(f"    Будет записано в Excel: {repr(row_values_num[quantity_column_pos])}")
    print(f"    Тип: {type(row_values_num[quantity_column_pos])}")

    # Имитация результата _record_to_row для record_formula_str
    row_values_formula_str = exporter._record_to_row(record_formula_str)
    print(f"  Для Строки-формулы Quantity ({record_formula_str.quantity}):")
    print(f"    Будет записано в Excel: {repr(row_values_formula_str[quantity_column_pos])}")
    print(f"    Тип: {type(row_values_formula_str[quantity_column_pos])}")

    # Имитация результата _record_to_row для record_text_str
    row_values_text_str = exporter._record_to_row(record_text_str)
    print(f"  Для Строки-текста Quantity ({record_text_str.quantity}):")
    print(f"    Будет записано в Excel: {repr(row_values_text_str[quantity_column_pos])}")
    print(f"    Тип: {type(row_values_text_str[quantity_column_pos])}")

    print("\nВывод: Если в ProductRecord.quantity попадает строка, начинающаяся с '=',")
    print("       ExcelExporter записывает ее как есть, и Excel интерпретирует это как формулу.")
    print("       Если 'data_only=False' при чтении, а формула не имеет кэшированного значения или")
    print("       зависит от внешних ссылок, Excel может выдать #ЗНАЧ! при открытии.")

if __name__ == "__main__":
    diagnose_exporter_quantity()
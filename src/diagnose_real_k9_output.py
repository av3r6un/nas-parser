from pathlib import Path
import sys

# Добавляем корневую директорию проекта в sys.path, чтобы можно было импортировать модули проекта
# Это временное изменение sys.path только для этого скрипта и не влияет на проект.
project_root = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(project_root))

from nas_parser.readers.excel import ExcelReader
from nas_parser.parsers.k9 import K9Parser
from nas_parser.domain import ProductRecord # Для type hints, не обязательно, но хорошая практика

FILE_PATH = Path("input/K9.xlsx")

def diagnose_real_k9_output():
    if not FILE_PATH.exists():
        print(f"Ошибка: Файл {FILE_PATH} не найден.")
        return

    print(f"--- Диагностика реальных ProductRecord из {FILE_PATH} после K9Parser ---")
    print("-" * 50)

    try:
        reader = ExcelReader(FILE_PATH)
        parser = K9Parser()
        
        # Обрабатываем записи и берем первые 20
        records = list(parser.parse(reader.read()))[:20]

        if not records:
            print("Не найдено ProductRecord для анализа.")
            return

        for r in records:
            print(f"Строка источника: {r.source_row}")
            print(f"  quantity: {repr(r.quantity)}, тип: {type(r.quantity)}")
            print(f"  price: {repr(r.price)}, тип: {type(r.price)}")
            print("-" * 20)

    except Exception as e:
        print(f"Произошла ошибка при обработке файла: {e}")

    print("-" * 50)

if __name__ == "__main__":
    diagnose_real_k9_output()
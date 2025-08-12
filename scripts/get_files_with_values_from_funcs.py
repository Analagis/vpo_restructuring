import argparse
from pathlib import Path
from openpyxl import load_workbook
import win32com.client as win32


def run_all_files_in_excel(files):
    """
    Один раз запускает Excel и сохраняет все файлы для пересчёта формул.
    """
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    excel.Visible = False
    excel.DisplayAlerts = False

    try:
        for file in files:
            print(f"[Excel] Открываю {file.name}")
            wb = excel.Workbooks.Open(str(file.resolve()))
            try:
                # Просто сохраняем — Excel при открытии пересчитает формулы
                wb.Save()
            finally:
                wb.Close(SaveChanges=True)
    finally:
        excel.Quit()


def replace_formulas_with_values(file_path: Path, output_path: Path):
    """
    Заменяет формулы на вычисленные значения.
    """
    wb = load_workbook(file_path, data_only=True)  # Берём сохранённые значения
    wb.save(output_path)


def process_folder(input_dir: Path):
    """
    Сначала прогоняет все файлы через Excel, потом делает версии без формул.
    """
    if not input_dir.is_dir():
        raise ValueError(f"{input_dir} не является папкой")

    output_dir = input_dir.parent / f"{input_dir.name}_values"
    output_dir.mkdir(exist_ok=True)

    # Получаем список файлов
    files = [f for f in input_dir.glob("*.xls*") if not f.name.startswith("~$")]

    # 1. Один запуск Excel для всех файлов
    run_all_files_in_excel(files)

    # 2. Проход через openpyxl
    for file in files:
        output_file = output_dir / file.name
        replace_formulas_with_values(file, output_file)
        print(f"[OK] Сохранён: {output_file}")

    print(f"[DONE] Все файлы сохранены в {output_dir}")


def process_file(input_file: Path):
    """
    Обрабатывает один файл (тоже через единый Excel-проход).
    """
    if not input_file.is_file():
        raise ValueError(f"{input_file} не найден")

    parent_dir = input_file.parent
    output_dir = parent_dir.parent / f"{parent_dir.name}_values"
    output_dir.mkdir(exist_ok=True)

    # 1. Запускаем Excel для одного файла
    run_all_files_in_excel([input_file])

    # 2. Через openpyxl убираем формулы
    output_file = output_dir / input_file.name
    replace_formulas_with_values(input_file, output_file)
    print(f"[OK] Файл сохранён: {output_file}")


def main():
    parser = argparse.ArgumentParser(description="Пересчитывает формулы в Excel и заменяет их на значения")
    parser.add_argument("path", type=str, help="Путь к файлу или папке с Excel-файлами")
    args = parser.parse_args()

    path = Path(args.path)
    if path.is_file():
        process_file(path)
    elif path.is_dir():
        process_folder(path)
    else:
        raise FileNotFoundError(f"Путь не найден: {path}")


if __name__ == "__main__":
    main()

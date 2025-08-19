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


def process_folder(input_dir: Path, mode: str):
    """
    Обрабатывает папку: если mode=True — делает только прогонку файлов,
    иначе — дополнительно сохраняет версии без формул.
    """
    if not input_dir.is_dir():
        raise ValueError(f"{input_dir} не является папкой")

    output_dir = input_dir.parent / f"{input_dir.name}_values"

    # Получаем список файлов
    files = [f for f in input_dir.glob("*.xls*") if not f.name.startswith("~$")]

    if mode in ("full", "recalc_only"):
        # 1. Один запуск Excel для всех файлов (пересчёт формул)
        run_all_files_in_excel(files)

    if mode in ("full", "values_only"):
        # 2. Через openpyxl сохраняем версии без формул
        output_dir.mkdir(exist_ok=True)
        for file in files:
            output_file = output_dir / file.name
            replace_formulas_with_values(file, output_file)
            print(f"[OK] Сохранён: {output_file}")

        print(f"[DONE] Все файлы сохранены в {output_dir}")

    print(f"[DONE] Режим '{mode}' завершён")

def process_file(input_file: Path, mode: str):
    """
    Обрабатывает один файл: если mode=True — только пересчёт,
    иначе — сохраняет версию без формул.
    """
    if not input_file.is_file():
        raise ValueError(f"{input_file} не найден")

    parent_dir = input_file.parent
    output_dir = parent_dir.parent / f"{parent_dir.name}_values"

    if mode in ("full", "recalc_only"):
        # 1. Пересчёт формул через Excel
        run_all_files_in_excel([input_file])

    if mode in ("full", "values_only"):
        # 2. Через openpyxl сохраняем версию без формул
        output_dir.mkdir(exist_ok=True)
        output_file = output_dir / input_file.name
        replace_formulas_with_values(input_file, output_file)
        print(f"[OK] Файл сохранён: {output_file}")

    print(f"[DONE] Режим '{mode}' завершён для файла {input_file.name}")


def main():
    parser = argparse.ArgumentParser(description="Пересчитывает формулы в Excel и заменяет их на значения")
    parser.add_argument("path", type=str, help="Путь к файлу или папке с Excel-файлами")
    parser.add_argument(
        "--mode", choices=["full", "recalc_only", "values_only"], default="full",
        help="Режим работы: 'full' - пересчёт и сохранение без формул; 'recalc_only' - только пересчёт и сохранение файлов (без замены формул); 'values_only' — сохраняет значения из уже пересчитанных файлов без открытия Excel"
    )
    args = parser.parse_args()

    path = Path(args.path)

    if path.is_file():
        process_file(path, args.mode)
    elif path.is_dir():
        process_folder(path, args.mode)
    else:
        raise FileNotFoundError(f"Путь не найден: {path}")


if __name__ == "__main__":
    main()

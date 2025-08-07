import os
import commentjson as json
import zipfile
from pathlib import Path
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter, column_index_from_string
from typing import Union, List, Dict, Optional
import psutil
import time
from contextlib import contextmanager
import win32com.client as win32
import re

class ExcelProcessor:
    def __init__(self, config_path: str|Path, template_path: Optional[str|Path] = None, output_dir: Optional[str|Path] = None) -> None:
        """
        Инициализация процессора с использованием xlwings
        :param config_path: Путь к файлу конфигурации
        :param template_path: Путь к шаблону (опционально)
        :param output_dir: Папка для сохранения результатов (опционально)
        """
        # Загрузка конфигурации
        with open(config_path, 'r', encoding='utf-8') as f:
            self.config = json.load(f)
        
        self.common = self.config["common"]
        self.years_data = {k: v for k, v in self.config.items() if k != "common"}
        
        # Установка путей
        self.template_path = template_path or os.path.join(
            os.path.dirname(__file__), "templates", "template.xlsx"
        )
        self.output_dir = output_dir or os.path.join(
            os.path.dirname(os.path.dirname(__file__)), "created_files"
        )
        
        # Создание папки для результатов
        os.makedirs(self.output_dir, exist_ok=True)

    def process_years(self, input_dir: str|Path, selected_years: Union[List[str], str] = "all", include_optional: bool = False) -> None:
        """
        Основной метод обработки данных
        :param input_dir: Папка с исходными данными
        :param selected_years: Список годов или "all"
        :param include_optional: Включать optional-файлы
        """
        years_to_process = self._get_years_to_process(selected_years)

        try:
            for year in years_to_process:
                year_dir = os.path.join(input_dir, "VPO_1_"+str(year))
                print(f"Обработка года: {year}, название папки/архива: {year_dir}")
                
                files = self._find_files(year_dir, include_optional)
                print(f"Найдено файлов для обработки: {len(files)}")

                output_path = self._get_output_path(year)

                # 1. Загружаем шаблон
                template_wb = self._load_template_workbook()

                # 2. Создаём новую книгу с листами из шаблона
                new_wb = self._create_new_workbook_with_template_sheets(template_wb)
                
                for file_path in files:
                    # 3. Обрабатываем данные и заполняем листы
                    data_columns = self._process_data_for_year(new_wb, file_path, year)
                    # 4. Создаём итоговый лист
                    self._create_summary_sheet(new_wb, data_columns)

                # 5. Удаляем лист-шаблон
                self._remove_sheet(new_wb)
                # 5. Сохраняем
                self._save_workbook(new_wb, output_path)
        except Exception as e:
            print(f"Критическая ошибка при обработке файла: {str(e)}")
            raise

    def _get_years_to_process(self, selected_years: Union[List[str], str]) -> List[str]:
        """Определяет какие годы нужно обработать"""
        if selected_years == "all":
            return list(self.years_data.keys())
        elif isinstance(selected_years, (list, tuple)):
            return [year for year in selected_years if year in self.years_data]
        return [selected_years] if selected_years in self.years_data else []

    def _find_files(self, path: str, include_optional: bool) -> List[str]:
        """Находит файлы для обработки с учетом фильтров"""
        zip_path = path+".zip"
        if zipfile.is_zipfile(zip_path):
            with zipfile.ZipFile(zip_path) as z:
                matched_files = []
                for zip_info in z.infolist():
                    if zip_info.is_dir():
                        continue
                    
                    if self._match_patterns(zip_info.filename, include_optional):
                        extracted_path = z.extract(zip_info, path=os.path.dirname(zip_path))
                        matched_files.append(extracted_path)
                
                return matched_files
        else:
            return [
                os.path.join(root, name)
                for root, _, files in os.walk(path)
                for name in files
                if self._match_patterns(name, include_optional)
            ]

    def _match_patterns(self, filename: str, include_optional: bool) -> bool:
        """Проверяет соответствие файла шаблонам имен"""
        required_patterns = [p.lower() for p in self.common["search_patterns"]["required"]]
        if not include_optional:
            required_ok = any(p in filename.lower()
                for p in required_patterns
            )
            return required_ok
        else:
            optional_patterns = [p.lower() for p in self.common["search_patterns"]["optional"]]
            optional_ok = any((p in filename.lower() and r in filename.lower())
                for p in optional_patterns
                for r in required_patterns
            )
            
            return optional_ok

    def _get_output_path(self, year: str) -> str:
        """Генерирует путь для выходного файла"""
        source_name = Path(self.template_path).stem.split("_")[:-1]
        output_name = f"{year}_{"_".join(source_name)}.xlsx"
        return os.path.join(self.output_dir, output_name)

    def _find_first_empty_column(self, sheet) -> str:
        """Находит первый пустой столбец в первой строке"""
        col = 1
        while sheet.cell(row=1, column=col).value is not None:
            col += 1
        return get_column_letter(col)
    
    def _copy_sheet_structure(self, src_sheet, dst_sheet) -> None:
        """Копирует данные и ширину столбцов"""
        # Копируем данные
        for row in src_sheet.iter_rows():
            for cell in row:
                dst_cell = dst_sheet[cell.coordinate]
                dst_cell.value = cell.value

        # Копируем ширину столбцов
        for col_idx in range(1, src_sheet.max_column + 1):
            col_letter = get_column_letter(col_idx)
            if src_sheet.column_dimensions[col_letter].width:
                dst_sheet.column_dimensions[col_letter].width = src_sheet.column_dimensions[col_letter].width


    def _fill_sheet(self, sheet, formulas: Dict, level_code: str, source_filename: str) -> str:
        """Заполняет лист данными и формулами"""
        # Находим первый пустой столбец
        start_col = self._find_first_empty_column(sheet)
        start_col_idx = column_index_from_string(start_col)

        # Заголовок — имя файла без расширения
        sheet[f"{start_col}1"].value = Path(source_filename).stem + "_" + self._get_data_type_by_sheet_name(formulas["params"]["sheet_name"])
        
        # Извлекаем саму формулу
        if not isinstance(formulas, dict) or "formula" not in formulas:
            raise ValueError("formulas должен быть словарём с ключом 'formula'")

        base_formula = formulas["formula"]  # строка вида "=INDEX(..., A2, ...)"
        last_row = sheet.max_row
        
        # Вставляем формулу во все строки, начиная со 2-й
        for row in range(2, last_row + 1):
            cell = sheet[f"{start_col}{row}"]
            cell.value = base_formula.replace(f"={formulas["params"]["row_condition"]}", f"={formulas["params"]["row_condition"][:-1]}{row}")
            
        print(base_formula)
        return start_col
    
        

    def _create_summary_sheet(self, wb, data_columns: Dict) -> None:
        """Создает итоговый лист"""
        if "Итоги" in wb.sheetnames:
            sheet = wb["Итоги"]
        else:
            sheet = wb.create_sheet("Итоги")
            if wb.sheetnames[0] != "Итоги":
                self._copy_sheet_structure(wb.worksheets[0], sheet)

        start_col = self._find_first_empty_column(sheet)
        
        for sheet_name in wb.sheetnames:
            if sheet_name in self.common["education"]["output_sheets"].values():
                try:
                    header = wb[sheet_name][f"{start_col}1"].value
                    if header:
                        sheet[f"{start_col}1"].value = f"Итого {header}"
                        break
                except:
                    continue
        else:
            sheet[f"{start_col}1"].value = "Итого"

        last_row = sheet.max_row
        for row in range(2, last_row + 1):
            sum_parts = []
            for level_sheet_name in self.common["education"]["output_sheets"].values():
                if level_sheet_name in wb.sheetnames:
                    sum_parts.append(f"'{level_sheet_name}'!{start_col}{row}")

            if sum_parts:
                sheet[f"{start_col}{row}"].value = f"=SUM({','.join(sum_parts)})"

    def _load_template_workbook(self):
        """Загружает шаблон"""
        return load_workbook(self.template_path)

    def _create_new_workbook_with_template_sheets(self, template_wb):
        """Создаёт новую книгу и копирует листы из шаблона"""
        new_wb = Workbook()
        # Удаляем стандартный лист
        default_sheet = new_wb.active
        new_wb.remove(default_sheet)

        for sheet in template_wb.worksheets:
            new_sheet = new_wb.create_sheet(sheet.title)
            self._copy_sheet_structure(sheet, new_sheet)

        return new_wb

    def _nan_error_handle_(self, base_formula: str, 
                     handle_nan: bool = True, 
                     handle_errors: bool = False,
                     nan_replacement: str|int = 0,
                     error_replacement: str|int = "#N/A") -> str:
        formula = base_formula
    
        # Обработка NaN
        if handle_nan:
            formula = f"IFNA({formula}, {nan_replacement})"
        
        # Обработка ошибок (поверх обработки NaN)
        if handle_errors:
            formula = f"IFERROR({formula}, {error_replacement})"
        
        return formula

    def _generate_formulas(self, source_file_path: str, sheet_name: str, params: Dict, recipient_file_path: str, level_code: int) -> Dict:
        """
        Генерирует формулу с фильтрацией по level_code (например, B=01)
        """
        # Получаем относительный путь
        abs_source = os.path.abspath(source_file_path)
        abs_recipient_dir = os.path.dirname(os.path.abspath(recipient_file_path))
        rel_path = os.path.relpath(os.path.dirname(abs_source), abs_recipient_dir).replace('\\', '/')

        filename = os.path.basename(abs_source)

        # Формируем ссылку на лист
        if rel_path == ".":
            sheet_ref = f"[{filename}]{sheet_name}"
        else:
            sheet_ref = f"{rel_path}/[{filename}]{sheet_name}"
        full_ref = f"'{sheet_ref}'!"

        array = params["array"]  # например, $A$12:$W$467
        row_params = params["row_num"]
        col_params = params["columns_num"]

        # Извлекаем границы диапазона
        arr_parts = array.split(':')
        start_row = int(arr_parts[0].split('$')[-1])  # 12 из $A$12
        start_col_letter = arr_parts[0].split('$')[1]  # $A$12 → A
        end_col_letter = arr_parts[1].split('$')[1]    # $W$467 → W
        end_row = int(arr_parts[1].split('$')[-1])

        # Диапазоны
        key_col = row_params['looup_array']  # например, D
        key_col_range = f"{full_ref}${key_col}${start_row}:${key_col}${end_row}"  # $D$12:$D$467

        edu_col = self.common["education"]["column"]  # B
        edu_col_range = f"{full_ref}${edu_col}${start_row}:${edu_col}${end_row}"  # $B$12:$B$467

        header_row_range = f"{full_ref}${start_col_letter}${start_row}:${end_col_letter}${start_row}"  # $A$12:$W$12

        # AGGREGATE для поиска строки с учётом level_code
        aggregate_part = (
            f"MATCH(1,"
            f"({key_col_range}={row_params['lookup_value']})*"
            f"({edu_col_range}={level_code}),"
            f"0)"
        )

        # Номер столбца
        col_match = f"MATCH({col_params['lookup_value']},{header_row_range},{col_params['match_type']})"

        # Финальная формула
        formula = f"INDEX({full_ref}{array},{aggregate_part},{col_match})"

        params["full_ref"] = full_ref
        params["row_condition"] = row_params['lookup_value']
        params["sheet_name"] = sheet_name
        return {
            "formula": f"={self._nan_error_handle_(formula)}",
            "params": params
        }

    def _process_data_for_year(self, new_wb, source_file_path: str, year: str) -> Dict:
        """Обрабатывает данные и заполняет листы"""
        year_data = self.years_data[year]
        data_columns = {}
        source_filename = Path(source_file_path).name

        for data_type, params in year_data["funcs"].items():
            sheet_name = self.common["cols_aliases"].get(data_type, data_type)

            for level, level_code in self.common["education"]["codes"].items():
                output_sheet_name = self.common["education"]["output_sheets"][level]

                # Генерируем формулу ДЛЯ ЭТОГО УРОВНЯ
                formulas = self._generate_formulas(
                    source_file_path=source_file_path,
                    sheet_name=sheet_name,
                    params=params,
                    recipient_file_path=self._get_output_path(year),
                    level_code=level_code  # ← передаём код уровня
                )

                # Ищем или создаём лист
                if output_sheet_name not in new_wb.sheetnames:
                    base_sheet = new_wb.worksheets[0]
                    new_sheet = new_wb.create_sheet(output_sheet_name)
                    self._copy_sheet_structure(base_sheet, new_sheet)
                else:
                    new_sheet = new_wb[output_sheet_name]

                # Заполняем
                try:
                    start_col = self._fill_sheet(new_sheet, formulas, level_code, source_filename)
                    data_columns[output_sheet_name] = start_col
                except Exception as e:
                    print(f"Ошибка заполнения листа {output_sheet_name}: {str(e)}")

        return data_columns

    def _remove_sheet(self, wb, sheet_reference: Union[str, int] = 0) -> None:
        """
        Удаляет лист из рабочей книги по названию или индексу
        
        :param wb: Рабочая книга (openpyxl.workbook.Workbook)
        :param sheet_reference: Название листа или его индекс (начиная с 0)
        """
        try:
            if isinstance(sheet_reference, str):
                # Удаление по названию
                if sheet_reference in wb.sheetnames:
                    sheet = wb[sheet_reference]
                    wb.remove(sheet)
            elif isinstance(sheet_reference, int):
                # Удаление по индексу
                if 0 <= sheet_reference < len(wb.sheetnames):
                    sheet = wb.worksheets[sheet_reference]
                    wb.remove(sheet)
        except Exception as e:
            print(f"Ошибка при удалении листа: {str(e)}")

    def _save_workbook(self, wb, output_path: str):
        """Сохраняет книгу"""
        try:
            wb.save(output_path)
            print(f"Файл успешно сохранён: {output_path}")
        except Exception as e:
            print(f"Ошибка сохранения: {str(e)}")
            raise

    def _get_data_type_by_sheet_name(self, sheet_name: str) -> str:
        """Возвращает логический тип по имени листа"""
        reverse_map = {v: k for k, v in self.common["cols_aliases"].items()}
        return reverse_map.get(sheet_name, "unknown")

    def _delete_dogs(self, file_name: str):
        import subprocess

        ps_script = f"""
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $wb = $excel.Workbooks.Open("{file_name}")

        foreach ($sheet in $wb.Worksheets) {{
            $sheet.UsedRange.Replace("@", "", [Type]::Missing, [Type]::Missing, $false)
        }}

        $wb.Save()
        $wb.Close()
        $excel.Quit()
        """

        subprocess.run(["powershell", "-Command", ps_script], shell=True)

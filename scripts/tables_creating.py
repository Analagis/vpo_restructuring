from pathlib import Path
import argparse
import os
from typing import Optional, Union, List
from vpo_from_json import ExcelProcessor

class Config:
    def __init__(
        self,
        config_path: str,
        input_dir: str,
        template_path: Optional[str] = None,
        output_dir: Optional[str] = None,
        selected_years: Union[List[str], str] = "all",
        include_optional: bool = False
    ):
        self.config_path = Path(config_path).absolute()
        self.input_dir = Path(input_dir).absolute()
        self.template_path = Path(template_path).absolute() if template_path else None
        self.output_dir = Path(output_dir).absolute() if output_dir else None
        self.selected_years = selected_years
        self.include_optional = include_optional

        self._validate_paths()

    def _validate_paths(self):
        if not self.config_path.exists():
            raise FileNotFoundError(f"Config file not found: {self.config_path}")
        
        if not self.input_dir.exists():
            raise FileNotFoundError(f"Input directory not found: {self.input_dir}")
        
        if self.template_path and not self.template_path.exists():
            raise FileNotFoundError(f"Template file not found: {self.template_path}")

class App:
    @staticmethod
    def parse_args() -> Config:
        parser = argparse.ArgumentParser(description='Создание таблиц из Excel-файлов')
        
        # Обязательные аргументы
        parser.add_argument(
            'config_path',
            type=str,
            help='Путь к JSON-файлу с конфигурацией'
        )
        parser.add_argument(
            'input_dir',
            type=str,
            help='Папка с исходными данными (по годам)'
        )
        
        # Опциональные аргументы
        parser.add_argument(
            '--template',
            dest='template_path',
            type=str,
            default=None,
            help='Путь к файлу шаблона (если не указан, используется стандартный)'
        )
        parser.add_argument(
            '--output',
            dest='output_dir',
            type=str,
            default=None,
            help='Папка для сохранения результатов (по умолчанию created_files)'
        )
        parser.add_argument(
            '--years',
            nargs='+',
            default="all",
            help='Годы для обработки (через пробел) или "all" для всех'
        )
        parser.add_argument(
            '--include_optional',
            action='store_true',
            help='Включать файлы с optional-шаблонами в названии'
        )
        
        args = parser.parse_args()
        
        return Config(
            config_path=args.config_path,
            input_dir=args.input_dir,
            template_path=args.template_path,
            output_dir=args.output_dir,
            selected_years=args.years,
            include_optional=args.include_optional
        )

    @classmethod
    def run(cls) -> None:
        try:
            config = cls.parse_args()
            
            processor = ExcelProcessor(
                config_path=config.config_path,
                template_path=config.template_path,
                output_dir=config.output_dir
            )
            
            processor.process_years(
                input_dir=config.input_dir,
                selected_years=config.selected_years,
                include_optional=config.include_optional
            )
            
            print("Обработка успешно завершена!")
            
        except Exception as e:
            print(f"Ошибка: {str(e)}")
            exit(1)

if __name__ == "__main__":
    print("here")
    App.run()
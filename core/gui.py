from __future__ import annotations

import queue
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from .core import RenameError, apply_rename_plan, build_rename_plan
from .xlsx_csv import ConversionCancelled, XlsxCsvError, convert_xlsx_path_to_csv, list_xlsx_sheets


class XlsxEditDependencyError(Exception):
    """XLSX editor cannot run until optional dependencies are installed."""


TRANSLATIONS = {
    "ru": {
        "Language": "Язык",
        "English": "Английский",
        "Russian": "Русский",
        "Batch rename": "Переименование",
        "XLSX to CSV": "XLSX в CSV",
        "XLSX editor": "Редактор XLSX",
        "CSV editor": "Редактор CSV",
        "Folder": "Папка",
        "Input mask": "Входная маска",
        "Output mask": "Выходная маска",
        "Recursive": "Рекурсивно",
        "Include folders": "Включать папки",
        "Overwrite existing": "Перезаписывать",
        "Preview": "Предпросмотр",
        "Apply rename": "Переименовать",
        "Current name": "Текущее имя",
        "New name": "Новое имя",
        "XLSX source": "Источник XLSX",
        "CSV source": "Источник CSV",
        "CSV target": "CSV назначение",
        "Choose file": "Выбрать файл",
        "Choose folder": "Выбрать папку",
        "Sheet": "Лист",
        "Load sheets": "Загрузить листы",
        "Delimiter": "Разделитель",
        "Semicolon": "Точка с запятой",
        "Comma": "Запятая",
        "Tab": "Табуляция",
        "Encoding": "Кодировка",
        "Recursive folder conversion": "Рекурсивная конвертация папки",
        "Convert": "Конвертировать",
        "Cancel": "Отмена",
        "Insert column": "Вставить колонку",
        "Delete column": "Удалить колонку",
        "Swap columns": "Поменять колонки",
        "Column": "Колонка",
        "Header": "Заголовок",
        "First column": "Первая колонка",
        "Second column": "Вторая колонка",
        "Recursive folder edit": "Рекурсивная обработка папки",
        "Recursive CSV folder edit": "Рекурсивная обработка папки CSV",
        "Check columns": "Проверить колонки",
        "Select all files": "Выбрать все файлы",
        "Use": "Исп.",
        "File": "Файл",
        "Headers": "Заголовки",
        "Columns": "Колонки",
        "Browse": "Обзор",
    }
}


class FileToolsApp(tk.Tk):
    # Класс окна остается тонкой оболочкой над backend: он собирает параметры
    # из виджетов, вызывает функции ядра и показывает пользователю результат.
    def __init__(self) -> None:
        super().__init__()
        self.title("Data Sculptor")
        self.set_app_icon()
        self.geometry("900x600")
        self.minsize(760, 500)

        self.rename_plans = []
        self.xlsx_sheets: list[str] = []
        self.conversion_queue: queue.Queue[tuple[str, object]] = queue.Queue()
        self.conversion_running = False
        self.cancel_conversion_event = threading.Event()
        self.edit_queue: queue.Queue[tuple[str, object]] = queue.Queue()
        self.edit_running = False
        self.cancel_edit_event = threading.Event()
        self.edit_checked_sources: set[str] = set()
        self.xlsx_edit_backend = None
        self.csv_edit_queue: queue.Queue[tuple[str, object]] = queue.Queue()
        self.csv_edit_running = False
        self.cancel_csv_edit_event = threading.Event()
        self.csv_checked_sources: set[str] = set()
        self.language = tk.StringVar(value="en")
        self.localized_widgets: list[tuple[object, str]] = []
        self.localized_headings: list[tuple[ttk.Treeview, str, str]] = []
        self.localized_tabs: list[tuple[ttk.Frame, str]] = []

        self.build_menu()

        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        self.rename_tab = ttk.Frame(self.notebook)
        self.xlsx_tab = ttk.Frame(self.notebook)
        self.xlsx_edit_tab = ttk.Frame(self.notebook)
        self.csv_edit_tab = ttk.Frame(self.notebook)
        self.add_localized_tab(self.rename_tab, "Batch rename")
        self.add_localized_tab(self.xlsx_tab, "XLSX to CSV")
        self.add_localized_tab(self.xlsx_edit_tab, "XLSX editor")
        self.add_localized_tab(self.csv_edit_tab, "CSV editor")

        self.build_rename_tab()
        self.build_xlsx_tab()
        self.build_xlsx_edit_tab()
        self.build_csv_edit_tab()
        self.after(300, self.check_xlsx_editor_dependency_on_startup) #TODO похоже эта строка не актуальна, мы избавились от проверки зависимостей приложения, т.к. оно всё равно не работает. Убери эту строку и все зависимые функции

    def tr(self, text: str) -> str:
        return TRANSLATIONS.get(self.language.get(), {}).get(text, text)

    def build_menu(self) -> None:
        menu_bar = tk.Menu(self)
        language_menu = tk.Menu(menu_bar, tearoff=False)
        language_menu.add_radiobutton(
            label=self.tr("English"),
            variable=self.language,
            value="en",
            command=lambda: self.set_language("en"),
        )
        language_menu.add_radiobutton(
            label=self.tr("Russian"),
            variable=self.language,
            value="ru",
            command=lambda: self.set_language("ru"),
        )
        menu_bar.add_cascade(label=self.tr("Language"), menu=language_menu)
        self.config(menu=menu_bar)

    def set_language(self, language: str) -> None:
        self.language.set(language)
        self.build_menu()
        for widget, text in self.localized_widgets:
            widget.configure(text=self.tr(text))
        for tree, column, text in self.localized_headings:
            tree.heading(column, text=self.tr(text))
        for tab, text in self.localized_tabs:
            self.notebook.tab(tab, text=self.tr(text))

    def localize_widget(self, widget, text: str):
        widget.configure(text=self.tr(text))
        self.localized_widgets.append((widget, text))
        return widget

    def add_localized_tab(self, tab: ttk.Frame, text: str) -> None:
        self.notebook.add(tab, text=self.tr(text))
        self.localized_tabs.append((tab, text))

    def localize_heading(self, tree: ttk.Treeview, column: str, text: str) -> None:
        tree.heading(column, text=self.tr(text))
        self.localized_headings.append((tree, column, text))

    def set_app_icon(self) -> None:
        icon_path = Path(__file__).resolve().parent.parent / "IMG" / "icon.ico"
        if icon_path.exists():
            try:
                self.iconbitmap(default=str(icon_path))
            except tk.TclError:
                pass

    def build_rename_tab(self) -> None:
        # Вкладка переименования ничего не делает с файлами до кнопки Apply:
        # сначала строится план, затем пользователь видит таблицу source->target.
        self.rename_dir = tk.StringVar()
        self.input_mask = tk.StringVar(value="{name} {date:DD.MM.YYYY}.{ext}")
        self.output_mask = tk.StringVar(value="{date:YYYYMMDD}_{name}.{ext}")
        self.rename_recursive = tk.BooleanVar(value=False)
        self.rename_include_dirs = tk.BooleanVar(value=False)
        self.rename_overwrite = tk.BooleanVar(value=False)

        form = ttk.Frame(self.rename_tab)
        form.pack(fill=tk.X, padx=8, pady=8)
        form.columnconfigure(1, weight=1)

        self.add_entry_row(form, 0, "Folder", self.rename_dir, self.choose_rename_dir)
        self.add_entry_row(form, 1, "Input mask", self.input_mask)
        self.add_entry_row(form, 2, "Output mask", self.output_mask)

        options = ttk.Frame(form)
        options.grid(row=3, column=1, sticky="w", pady=(6, 0))
        self.localize_widget(ttk.Checkbutton(options, variable=self.rename_recursive), "Recursive").pack(side=tk.LEFT, padx=(0, 12))
        self.localize_widget(ttk.Checkbutton(options, variable=self.rename_include_dirs), "Include folders").pack(side=tk.LEFT, padx=(0, 12))
        self.localize_widget(ttk.Checkbutton(options, variable=self.rename_overwrite), "Overwrite existing").pack(side=tk.LEFT)

        actions = ttk.Frame(form)
        actions.grid(row=4, column=1, sticky="w", pady=(10, 0))
        self.localize_widget(ttk.Button(actions, command=self.preview_renames), "Preview").pack(side=tk.LEFT, padx=(0, 8))
        self.localize_widget(ttk.Button(actions, command=self.apply_renames), "Apply rename").pack(side=tk.LEFT)

        table_frame = ttk.Frame(self.rename_tab)
        table_frame.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0, 8))
        table_frame.columnconfigure(0, weight=1)
        table_frame.rowconfigure(0, weight=1)

        self.rename_table = ttk.Treeview(table_frame, columns=("source", "target"), show="headings")
        self.localize_heading(self.rename_table, "source", "Current name")
        self.localize_heading(self.rename_table, "target", "New name")
        self.rename_table.column("source", width=360)
        self.rename_table.column("target", width=360)
        self.rename_table.grid(row=0, column=0, sticky="nsew")

        scrollbar = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.rename_table.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.rename_table.configure(yscrollcommand=scrollbar.set)

        self.rename_status = tk.StringVar(value="Ready.")
        ttk.Label(self.rename_tab, textvariable=self.rename_status).pack(fill=tk.X, padx=8, pady=(0, 8))

    def build_xlsx_tab(self) -> None:
        # Вкладка XLSX поддерживает два режима: один файл или целая папка.
        # Одинаковые поля используются для обоих режимов, чтобы UI был компактным.
        self.xlsx_source = tk.StringVar()
        self.csv_target = tk.StringVar()
        self.csv_delimiter = tk.StringVar(value=";")
        self.xlsx_sheet = tk.StringVar()
        self.csv_encoding = tk.StringVar(value="utf-8-sig")
        self.xlsx_recursive = tk.BooleanVar(value=False)

        form = ttk.Frame(self.xlsx_tab)
        form.pack(fill=tk.X, padx=8, pady=8)
        form.columnconfigure(1, weight=1)

        self.add_multi_button_entry_row(
            form,
            0,
            "XLSX source",
            self.xlsx_source,
            (("Choose file", self.choose_xlsx_source_file), ("Choose folder", self.choose_xlsx_source_dir)),
        )
        self.add_multi_button_entry_row(
            form,
            1,
            "CSV target",
            self.csv_target,
            (("Choose file", self.choose_csv_target_file), ("Choose folder", self.choose_csv_target_dir)),
        )

        self.localize_widget(ttk.Label(form), "Sheet").grid(row=2, column=0, sticky="w", padx=(0, 8), pady=4)
        sheet_row = ttk.Frame(form)
        sheet_row.grid(row=2, column=1, sticky="ew", pady=4)
        sheet_row.columnconfigure(0, weight=1)
        self.sheet_combo = ttk.Combobox(sheet_row, textvariable=self.xlsx_sheet, values=self.xlsx_sheets)
        self.sheet_combo.grid(row=0, column=0, sticky="ew")
        self.localize_widget(ttk.Button(sheet_row, command=self.load_sheets), "Load sheets").grid(row=0, column=1, padx=(8, 0))

        self.localize_widget(ttk.Label(form), "Delimiter").grid(row=3, column=0, sticky="w", padx=(0, 8), pady=4)
        delimiter_row = ttk.Frame(form)
        delimiter_row.grid(row=3, column=1, sticky="w", pady=4)
        for label, value in (("Semicolon", ";"), ("Comma", ","), ("Tab", "\\t")):
            self.localize_widget(ttk.Radiobutton(delimiter_row, value=value, variable=self.csv_delimiter), label).pack(
                side=tk.LEFT, padx=(0, 12)
            )
        ttk.Entry(delimiter_row, textvariable=self.csv_delimiter, width=6).pack(side=tk.LEFT)

        self.add_entry_row(form, 4, "Encoding", self.csv_encoding)

        options = ttk.Frame(form)
        options.grid(row=5, column=1, sticky="w", pady=(6, 0))
        self.localize_widget(ttk.Checkbutton(options, variable=self.xlsx_recursive), "Recursive folder conversion").pack(side=tk.LEFT)

        actions = ttk.Frame(form)
        actions.grid(row=6, column=1, sticky="w", pady=(10, 0))
        self.convert_button = self.localize_widget(ttk.Button(actions, command=self.convert_xlsx), "Convert")
        self.convert_button.pack(side=tk.LEFT, padx=(0, 8))
        self.cancel_button = self.localize_widget(ttk.Button(actions, command=self.cancel_xlsx_conversion, state=tk.DISABLED), "Cancel")
        self.cancel_button.pack(side=tk.LEFT)

        progress_frame = ttk.Frame(self.xlsx_tab)
        progress_frame.pack(fill=tk.X, padx=8, pady=(8, 0))
        progress_frame.columnconfigure(0, weight=1)

        self.xlsx_current_file = tk.StringVar(value="Current file: -")
        self.xlsx_progress_text = tk.StringVar(value="Progress: 0%")
        ttk.Label(progress_frame, textvariable=self.xlsx_current_file).grid(row=0, column=0, sticky="ew")
        ttk.Label(progress_frame, textvariable=self.xlsx_progress_text).grid(row=0, column=1, sticky="e", padx=(8, 0))

        self.xlsx_progress = ttk.Progressbar(progress_frame, mode="determinate", maximum=100)
        self.xlsx_progress.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(4, 0))

        self.xlsx_status = tk.StringVar(value="Ready.")
        ttk.Label(self.xlsx_tab, textvariable=self.xlsx_status).pack(fill=tk.X, padx=8, pady=(8, 0))

    def build_xlsx_edit_tab(self) -> None:
        # Третья вкладка работает поверх xlsx_edit.py: можно массово проверить
        # заполненные колонки и вставить новую колонку в один файл или папку.
        self.edit_source = tk.StringVar()
        self.edit_sheet = tk.StringVar()
        self.edit_insert_column = tk.StringVar(value="A")
        self.edit_delete_column = tk.StringVar(value="A")
        self.edit_swap_first_column = tk.StringVar(value="A")
        self.edit_swap_second_column = tk.StringVar(value="B")
        self.edit_header = tk.StringVar()
        self.edit_recursive = tk.BooleanVar(value=False)
        self.edit_select_all = tk.BooleanVar(value=False)

        form = ttk.Frame(self.xlsx_edit_tab)
        form.pack(fill=tk.X, padx=8, pady=8)
        form.columnconfigure(1, weight=1)

        self.add_multi_button_entry_row(
            form,
            0,
            "XLSX source",
            self.edit_source,
            (("Choose file", self.choose_edit_source_file), ("Choose folder", self.choose_edit_source_dir)),
        )
        self.add_entry_row(form, 1, "Sheet", self.edit_sheet)

        insert_group = self.localize_widget(ttk.LabelFrame(form), "Insert column")
        insert_group.grid(row=2, column=1, columnspan=2, sticky="ew", pady=(8, 0))
        insert_group.columnconfigure(1, weight=1)
        self.add_entry_row(insert_group, 0, "Column", self.edit_insert_column)
        self.add_entry_row(insert_group, 1, "Header", self.edit_header)
        self.edit_insert_button = self.localize_widget(ttk.Button(insert_group, command=self.insert_xlsx_column), "Insert column")
        self.edit_insert_button.grid(row=2, column=1, sticky="w", pady=(6, 4))

        delete_group = self.localize_widget(ttk.LabelFrame(form), "Delete column")
        delete_group.grid(row=3, column=1, columnspan=2, sticky="ew", pady=(8, 0))
        delete_group.columnconfigure(1, weight=1)
        self.add_entry_row(delete_group, 0, "Column", self.edit_delete_column)
        self.edit_delete_button = self.localize_widget(ttk.Button(delete_group, command=self.delete_xlsx_column), "Delete column")
        self.edit_delete_button.grid(row=1, column=1, sticky="w", pady=(6, 4))

        swap_group = self.localize_widget(ttk.LabelFrame(form), "Swap columns")
        swap_group.grid(row=4, column=1, columnspan=2, sticky="ew", pady=(8, 0))
        swap_group.columnconfigure(1, weight=1)
        self.add_entry_row(swap_group, 0, "First column", self.edit_swap_first_column)
        self.add_entry_row(swap_group, 1, "Second column", self.edit_swap_second_column)
        self.edit_swap_button = self.localize_widget(ttk.Button(swap_group, command=self.swap_xlsx_columns), "Swap columns")
        self.edit_swap_button.grid(row=2, column=1, sticky="w", pady=(6, 4))

        options = ttk.Frame(form)
        options.grid(row=5, column=1, sticky="w", pady=(6, 0))
        self.localize_widget(ttk.Checkbutton(options, variable=self.edit_recursive), "Recursive folder edit").pack(side=tk.LEFT)

        actions = ttk.Frame(form)
        actions.grid(row=6, column=1, sticky="w", pady=(10, 0))
        self.edit_check_button = self.localize_widget(ttk.Button(actions, command=self.check_xlsx_columns), "Check columns")
        self.edit_check_button.pack(side=tk.LEFT, padx=(0, 8))
        self.edit_cancel_button = self.localize_widget(ttk.Button(actions, command=self.cancel_xlsx_edit, state=tk.DISABLED), "Cancel")
        self.edit_cancel_button.pack(side=tk.LEFT)
        select_actions = ttk.Frame(form)
        select_actions.grid(row=7, column=1, sticky="w", pady=(6, 0))
        self.localize_widget(ttk.Checkbutton(
            select_actions,
            variable=self.edit_select_all,
            command=self.toggle_all_edit_rows,
        ), "Select all files").pack(side=tk.LEFT, padx=(0, 12))

        table_frame = ttk.Frame(self.xlsx_edit_tab)
        table_frame.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0, 8))
        table_frame.columnconfigure(0, weight=1)
        table_frame.rowconfigure(0, weight=1)

        self.edit_table = ttk.Treeview(
            table_frame,
            columns=("selected", "file", "sheet", "count", "columns"),
            show="headings",
        )
        self.localize_heading(self.edit_table, "selected", "Use")
        self.localize_heading(self.edit_table, "file", "File")
        self.localize_heading(self.edit_table, "sheet", "Sheet")
        self.localize_heading(self.edit_table, "count", "Headers")
        self.localize_heading(self.edit_table, "columns", "Columns")
        self.edit_table.column("selected", width=42, minwidth=38, stretch=False, anchor=tk.CENTER)
        self.edit_table.column("file", width=190, minwidth=140)
        self.edit_table.column("sheet", width=85, minwidth=70, stretch=False)
        self.edit_table.column("count", width=92, minwidth=78, stretch=False, anchor=tk.CENTER)
        self.edit_table.column("columns", width=720, minwidth=420)
        self.edit_table.grid(row=0, column=0, sticky="nsew")

        vertical_scrollbar = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.edit_table.yview)
        vertical_scrollbar.grid(row=0, column=1, sticky="ns")
        horizontal_scrollbar = ttk.Scrollbar(table_frame, orient=tk.HORIZONTAL, command=self.edit_table.xview)
        horizontal_scrollbar.grid(row=1, column=0, sticky="ew")
        self.edit_table.configure(
            yscrollcommand=vertical_scrollbar.set,
            xscrollcommand=horizontal_scrollbar.set,
        )
        self.edit_table.bind("<ButtonRelease-1>", self.toggle_edit_row)

        progress_frame = ttk.Frame(self.xlsx_edit_tab)
        progress_frame.pack(fill=tk.X, padx=8, pady=(0, 8))
        progress_frame.columnconfigure(0, weight=1)
        self.edit_current_file = tk.StringVar(value="Current file: -")
        self.edit_progress_text = tk.StringVar(value="Progress: 0%")
        ttk.Label(progress_frame, textvariable=self.edit_current_file).grid(row=0, column=0, sticky="ew")
        ttk.Label(progress_frame, textvariable=self.edit_progress_text).grid(row=0, column=1, sticky="e", padx=(8, 0))
        self.edit_progress = ttk.Progressbar(progress_frame, mode="determinate", maximum=100)
        self.edit_progress.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(4, 0))

        self.edit_status = tk.StringVar(value="Ready.")
        ttk.Label(self.xlsx_edit_tab, textvariable=self.edit_status).pack(fill=tk.X, padx=8, pady=(0, 8))

    def build_csv_edit_tab(self) -> None:
        self.csv_edit_source = tk.StringVar()
        self.csv_edit_delimiter = tk.StringVar(value=";")
        self.csv_edit_encoding = tk.StringVar(value="utf-8-sig")
        self.csv_edit_insert_column = tk.StringVar(value="A")
        self.csv_edit_delete_column = tk.StringVar(value="A")
        self.csv_edit_swap_first_column = tk.StringVar(value="A")
        self.csv_edit_swap_second_column = tk.StringVar(value="B")
        self.csv_edit_header = tk.StringVar()
        self.csv_edit_recursive = tk.BooleanVar(value=False)
        self.csv_edit_select_all = tk.BooleanVar(value=False)

        form = ttk.Frame(self.csv_edit_tab)
        form.pack(fill=tk.X, padx=8, pady=8)
        form.columnconfigure(1, weight=1)

        self.add_multi_button_entry_row(
            form,
            0,
            "CSV source",
            self.csv_edit_source,
            (("Choose file", self.choose_csv_edit_source_file), ("Choose folder", self.choose_csv_edit_source_dir)),
        )
        self.add_entry_row(form, 1, "Delimiter", self.csv_edit_delimiter)
        self.add_entry_row(form, 2, "Encoding", self.csv_edit_encoding)

        insert_group = self.localize_widget(ttk.LabelFrame(form), "Insert column")
        insert_group.grid(row=3, column=1, columnspan=2, sticky="ew", pady=(8, 0))
        insert_group.columnconfigure(1, weight=1)
        self.add_entry_row(insert_group, 0, "Column", self.csv_edit_insert_column)
        self.add_entry_row(insert_group, 1, "Header", self.csv_edit_header)
        self.csv_edit_insert_button = self.localize_widget(ttk.Button(insert_group, command=self.insert_csv_column), "Insert column")
        self.csv_edit_insert_button.grid(row=2, column=1, sticky="w", pady=(6, 4))

        delete_group = self.localize_widget(ttk.LabelFrame(form), "Delete column")
        delete_group.grid(row=4, column=1, columnspan=2, sticky="ew", pady=(8, 0))
        delete_group.columnconfigure(1, weight=1)
        self.add_entry_row(delete_group, 0, "Column", self.csv_edit_delete_column)
        self.csv_edit_delete_button = self.localize_widget(ttk.Button(delete_group, command=self.delete_csv_column), "Delete column")
        self.csv_edit_delete_button.grid(row=1, column=1, sticky="w", pady=(6, 4))

        swap_group = self.localize_widget(ttk.LabelFrame(form), "Swap columns")
        swap_group.grid(row=5, column=1, columnspan=2, sticky="ew", pady=(8, 0))
        swap_group.columnconfigure(1, weight=1)
        self.add_entry_row(swap_group, 0, "First column", self.csv_edit_swap_first_column)
        self.add_entry_row(swap_group, 1, "Second column", self.csv_edit_swap_second_column)
        self.csv_edit_swap_button = self.localize_widget(ttk.Button(swap_group, command=self.swap_csv_columns), "Swap columns")
        self.csv_edit_swap_button.grid(row=2, column=1, sticky="w", pady=(6, 4))

        options = ttk.Frame(form)
        options.grid(row=6, column=1, sticky="w", pady=(6, 0))
        self.localize_widget(ttk.Checkbutton(options, variable=self.csv_edit_recursive), "Recursive CSV folder edit").pack(side=tk.LEFT)

        actions = ttk.Frame(form)
        actions.grid(row=7, column=1, sticky="w", pady=(10, 0))
        self.csv_edit_check_button = self.localize_widget(ttk.Button(actions, command=self.check_csv_columns), "Check columns")
        self.csv_edit_check_button.pack(side=tk.LEFT, padx=(0, 8))
        self.csv_edit_cancel_button = self.localize_widget(ttk.Button(actions, command=self.cancel_csv_edit, state=tk.DISABLED), "Cancel")
        self.csv_edit_cancel_button.pack(side=tk.LEFT)
        select_actions = ttk.Frame(form)
        select_actions.grid(row=8, column=1, sticky="w", pady=(6, 0))
        self.localize_widget(ttk.Checkbutton(select_actions, variable=self.csv_edit_select_all, command=self.toggle_all_csv_edit_rows), "Select all files").pack(side=tk.LEFT)

        table_frame = ttk.Frame(self.csv_edit_tab)
        table_frame.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0, 8))
        table_frame.columnconfigure(0, weight=1)
        table_frame.rowconfigure(0, weight=1)
        self.csv_edit_table = ttk.Treeview(table_frame, columns=("selected", "file", "count", "columns"), show="headings")
        self.localize_heading(self.csv_edit_table, "selected", "Use")
        self.localize_heading(self.csv_edit_table, "file", "File")
        self.localize_heading(self.csv_edit_table, "count", "Headers")
        self.localize_heading(self.csv_edit_table, "columns", "Columns")
        self.csv_edit_table.column("selected", width=42, minwidth=38, stretch=False, anchor=tk.CENTER)
        self.csv_edit_table.column("file", width=220, minwidth=140)
        self.csv_edit_table.column("count", width=92, minwidth=78, stretch=False, anchor=tk.CENTER)
        self.csv_edit_table.column("columns", width=720, minwidth=420)
        self.csv_edit_table.grid(row=0, column=0, sticky="nsew")
        vertical_scrollbar = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.csv_edit_table.yview)
        vertical_scrollbar.grid(row=0, column=1, sticky="ns")
        horizontal_scrollbar = ttk.Scrollbar(table_frame, orient=tk.HORIZONTAL, command=self.csv_edit_table.xview)
        horizontal_scrollbar.grid(row=1, column=0, sticky="ew")
        self.csv_edit_table.configure(yscrollcommand=vertical_scrollbar.set, xscrollcommand=horizontal_scrollbar.set)
        self.csv_edit_table.bind("<ButtonRelease-1>", self.toggle_csv_edit_row)

        progress_frame = ttk.Frame(self.csv_edit_tab)
        progress_frame.pack(fill=tk.X, padx=8, pady=(0, 8))
        progress_frame.columnconfigure(0, weight=1)
        self.csv_edit_current_file = tk.StringVar(value="Current file: -")
        self.csv_edit_progress_text = tk.StringVar(value="Progress: 0%")
        ttk.Label(progress_frame, textvariable=self.csv_edit_current_file).grid(row=0, column=0, sticky="ew")
        ttk.Label(progress_frame, textvariable=self.csv_edit_progress_text).grid(row=0, column=1, sticky="e", padx=(8, 0))
        self.csv_edit_progress = ttk.Progressbar(progress_frame, mode="determinate", maximum=100)
        self.csv_edit_progress.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(4, 0))
        self.csv_edit_status = tk.StringVar(value="Ready.")
        ttk.Label(self.csv_edit_tab, textvariable=self.csv_edit_status).pack(fill=tk.X, padx=8, pady=(0, 8))

    def add_entry_row(
        self,
        parent: ttk.Frame,
        row: int,
        label: str,
        variable: tk.StringVar,
        command: object | None = None,
    ) -> None:
        # Универсальная строка формы: подпись, поле ввода и необязательная
        # кнопка выбора пути.
        self.localize_widget(ttk.Label(parent), label).grid(row=row, column=0, sticky="w", padx=(0, 8), pady=4)
        ttk.Entry(parent, textvariable=variable).grid(row=row, column=1, sticky="ew", pady=4)
        if command is not None:
            self.localize_widget(ttk.Button(parent, command=command), "Browse").grid(row=row, column=2, padx=(8, 0), pady=4)

    def add_multi_button_entry_row(
        self,
        parent: ttk.Frame,
        row: int,
        label: str,
        variable: tk.StringVar,
        buttons: tuple[tuple[str, object], ...],
    ) -> None:
        # Для XLSX нужны две кнопки выбора: файл и папка. Этот helper оставляет
        # сетку формы одинаковой и не размазывает layout-код по методам.
        self.localize_widget(ttk.Label(parent), label).grid(row=row, column=0, sticky="w", padx=(0, 8), pady=4)
        ttk.Entry(parent, textvariable=variable).grid(row=row, column=1, sticky="ew", pady=4)
        button_row = ttk.Frame(parent)
        button_row.grid(row=row, column=2, padx=(8, 0), pady=4)
        for text, command in buttons:
            self.localize_widget(ttk.Button(button_row, command=command), text).pack(side=tk.LEFT, padx=(0, 4))

    def choose_rename_dir(self) -> None:
        directory = filedialog.askdirectory(title="Choose folder")
        if directory:
            self.rename_dir.set(directory)

    def choose_xlsx_source_file(self) -> None:
        filename = filedialog.askopenfilename(
            title="Choose XLSX file",
            filetypes=[("Excel workbooks", "*.xlsx"), ("All files", "*.*")],
        )
        if filename:
            self.xlsx_source.set(filename)
            if not self.csv_target.get():
                self.csv_target.set(str(Path(filename).with_suffix(".csv")))
            self.load_sheets()

    def choose_xlsx_source_dir(self) -> None:
        directory = filedialog.askdirectory(title="Choose folder with XLSX files")
        if directory:
            self.xlsx_source.set(directory)
            self.xlsx_sheet.set("")
            self.xlsx_sheets = []
            self.sheet_combo.configure(values=[])
            if not self.csv_target.get() or Path(self.csv_target.get()).suffix.lower() == ".csv":
                # По умолчанию кладем CSV рядом с выбранной папкой в подпапку csv.
                self.csv_target.set(str(Path(directory) / "csv"))
            self.xlsx_status.set("Folder selected. Sheet field can be blank or a sheet name/index used in every file.")

    def choose_csv_target_file(self) -> None:
        filename = filedialog.asksaveasfilename(
            title="Choose CSV file",
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
        )
        if filename:
            self.csv_target.set(filename)

    def choose_csv_target_dir(self) -> None:
        directory = filedialog.askdirectory(title="Choose output folder")
        if directory:
            self.csv_target.set(directory)

    def choose_edit_source_file(self) -> None:
        filename = filedialog.askopenfilename(
            title="Choose XLSX file",
            filetypes=[("Excel workbooks", "*.xlsx"), ("All files", "*.*")],
        )
        if filename:
            self.edit_source.set(filename)

    def choose_edit_source_dir(self) -> None:
        directory = filedialog.askdirectory(title="Choose folder with XLSX files")
        if directory:
            self.edit_source.set(directory)

    def preview_renames(self) -> None:
        # Preview каждый раз пересчитывается заново, чтобы таблица отражала
        # текущие маски и состояние папки.
        for item in self.rename_table.get_children():
            self.rename_table.delete(item)

        try:
            self.rename_plans = build_rename_plan(
                Path(self.rename_dir.get()),
                self.input_mask.get(),
                self.output_mask.get(),
                recursive=self.rename_recursive.get(),
                include_dirs=self.rename_include_dirs.get(),
                overwrite=self.rename_overwrite.get(),
            )
        except RenameError as exc:
            self.rename_plans = []
            self.rename_status.set("Preview failed.")
            messagebox.showerror("Rename error", str(exc))
            return

        for plan in self.rename_plans:
            self.rename_table.insert("", tk.END, values=(plan.source.name, plan.target.name))

        self.rename_status.set(f"Preview: {len(self.rename_plans)} item(s).")

    def apply_renames(self) -> None:
        # Если пользователь сразу нажал Apply, сначала строим preview-план и
        # затем просим подтверждение уже с точным количеством объектов.
        if not self.rename_plans:
            self.preview_renames()
        if not self.rename_plans:
            return

        confirmed = messagebox.askyesno("Apply rename", f"Rename {len(self.rename_plans)} item(s)?")
        if not confirmed:
            return

        try:
            apply_rename_plan(self.rename_plans)
        except OSError as exc:
            self.rename_status.set("Rename failed.")
            messagebox.showerror("Rename error", str(exc))
            return

        count = len(self.rename_plans)
        self.rename_plans = []
        self.preview_renames()
        self.rename_status.set(f"Renamed: {count}")
        messagebox.showinfo("Batch rename", f"Completed.\nRenamed: {count} item(s).")

    def load_sheets(self) -> None:
        source = self.xlsx_source.get()
        if not source:
            return
        if Path(source).is_dir():
            self.xlsx_sheets = []
            self.sheet_combo.configure(values=[])
            self.xlsx_status.set("Folder selected. Sheet list is available only for a single XLSX file.")
            return

        try:
            self.xlsx_sheets = list_xlsx_sheets(Path(source))
        except XlsxCsvError as exc:
            self.xlsx_sheets = []
            self.sheet_combo.configure(values=[])
            self.xlsx_status.set("Could not load sheets.")
            messagebox.showerror("XLSX error", str(exc))
            return

        self.sheet_combo.configure(values=self.xlsx_sheets)
        if self.xlsx_sheets and not self.xlsx_sheet.get():
            self.xlsx_sheet.set(self.xlsx_sheets[0])
        self.xlsx_status.set(f"Loaded sheets: {len(self.xlsx_sheets)}")

    def convert_xlsx(self) -> None:
        if self.conversion_running:
            return
        if not self.xlsx_source.get() or not self.csv_target.get():
            messagebox.showerror("XLSX error", "Choose XLSX source and CSV target first.")
            return

        delimiter = self.csv_delimiter.get()
        delimiter = "\t" if delimiter == "\\t" else delimiter
        sheet = self.xlsx_sheet.get() or None
        source = Path(self.xlsx_source.get())
        target = Path(self.csv_target.get())
        encoding = self.csv_encoding.get()
        recursive = self.xlsx_recursive.get()

        # Перед стартом очищаем флаг отмены и старые события, иначе новый запуск
        # мог бы увидеть отмену от предыдущей операции.
        self.cancel_conversion_event.clear()
        self.drain_conversion_queue()
        self.conversion_running = True
        self.convert_button.configure(state=tk.DISABLED)
        self.cancel_button.configure(state=tk.NORMAL)
        self.xlsx_progress.configure(value=0)
        self.xlsx_current_file.set("Current file: preparing...")
        self.xlsx_progress_text.set("Progress: 0%")
        self.xlsx_status.set("Conversion started.")

        thread = threading.Thread(
            target=self.run_xlsx_conversion,
            args=(source, target, delimiter, sheet, encoding, recursive),
            daemon=True,
        )
        thread.start()
        self.after(100, self.process_conversion_queue)

    def run_xlsx_conversion(
        self,
        source: Path,
        target: Path,
        delimiter: str,
        sheet: str | None,
        encoding: str,
        recursive: bool,
    ) -> None:
        # Фоновый поток не трогает tkinter-виджеты напрямую. Все обновления GUI
        # передаются в очередь, а главный поток забирает их через after().
        try:
            results = convert_xlsx_path_to_csv(
                source,
                target,
                delimiter=delimiter,
                sheet=sheet,
                encoding=encoding,
                recursive=recursive,
                progress=lambda event: self.conversion_queue.put(("progress", event)),
                cancel_requested=self.cancel_conversion_event.is_set,
            )
        except ConversionCancelled:
            self.conversion_queue.put(("cancelled", None))
            return
        except XlsxCsvError as exc:
            self.conversion_queue.put(("error", f"XLSX error: {exc}"))
            return
        except OSError as exc:
            self.conversion_queue.put(("error", f"File error: {exc}"))
            return

        self.conversion_queue.put(("done", len(results)))

    def process_conversion_queue(self) -> None:
        # Этот метод всегда выполняется в главном tkinter-потоке. Благодаря
        # этому окно не зависает, а обновления прогресса безопасны для UI.
        while True:
            try:
                event_type, payload = self.conversion_queue.get_nowait()
            except queue.Empty:
                break

            if event_type == "progress":
                self.update_conversion_progress(payload)
            elif event_type == "error":
                self.finish_conversion()
                self.xlsx_status.set("Conversion failed.")
                messagebox.showerror("XLSX to CSV", str(payload))
                return
            elif event_type == "cancelled":
                self.finish_conversion()
                self.xlsx_status.set("Conversion cancelled.")
                self.xlsx_current_file.set("Current file: cancelled.")
                messagebox.showinfo("XLSX to CSV", "Conversion cancelled.")
                return
            elif event_type == "done":
                self.xlsx_progress.configure(value=100)
                self.xlsx_progress_text.set("Progress: 100%")
                self.finish_conversion()
                self.xlsx_status.set(f"Converted: {payload} file(s).")
                messagebox.showinfo("XLSX to CSV", f"Completed.\nConverted: {payload} file(s).")
                return

        if self.conversion_running:
            self.after(100, self.process_conversion_queue)

    def update_conversion_progress(self, event: object) -> None:
        # Progress bar показывает прогресс текущего файла, а счетчик рядом с
        # именем файла показывает положение во всей пачке.
        current_file = event.source.name
        self.xlsx_current_file.set(
            f"Current file: {event.file_index}/{event.file_count} - {current_file} ({event.message})"
        )
        self.xlsx_progress.configure(value=event.percent)
        self.xlsx_progress_text.set(f"Progress: {event.percent}%")
        self.xlsx_status.set(f"Converting: {current_file}")

    def finish_conversion(self) -> None:
        self.conversion_running = False
        self.convert_button.configure(state=tk.NORMAL)
        self.cancel_button.configure(state=tk.DISABLED)

    def cancel_xlsx_conversion(self) -> None:
        if not self.conversion_running:
            return
        self.cancel_conversion_event.set()
        self.cancel_button.configure(state=tk.DISABLED)
        self.xlsx_status.set("Cancel requested. Waiting for current stage to stop...")

    def drain_conversion_queue(self) -> None:
        while True:
            try:
                self.conversion_queue.get_nowait()
            except queue.Empty:
                return

    def check_xlsx_columns(self, *, show_done: bool = True) -> None:
        if self.edit_running:
            return
        if not self.edit_source.get():
            messagebox.showerror("XLSX editor", "Choose XLSX source first.")
            return

        try:
            files = self.collect_edit_source_files()
        except (XlsxEditDependencyError, ValueError) as exc:
            messagebox.showerror("XLSX editor", str(exc))
            return

        for item in self.edit_table.get_children():
            self.edit_table.delete(item)
        self.edit_checked_sources.clear()
        self.edit_select_all.set(False)
        self.start_xlsx_edit_operation("check", files, show_done=show_done)

    def populate_edit_table(self, reports: list[object]) -> None:
        for report in reports:
            columns_text = ", ".join(
                f"{column.letter}: {column.header or '-'}" for column in report.columns
            )
            source_key = str(report.source)
            self.edit_checked_sources.add(source_key)
            self.edit_table.insert(
                "",
                tk.END,
                iid=source_key,
                values=("[x]", report.source.name, report.sheet_name, report.filled_column_count, columns_text),
            )
        self.sync_edit_select_all_checkbox()

    def insert_xlsx_column(self) -> None:
        if self.edit_running:
            return
        if not self.edit_source.get():
            messagebox.showerror("XLSX editor", "Choose XLSX source first.")
            return
        if not self.edit_insert_column.get():
            messagebox.showerror("XLSX editor", "Enter insert column, for example A, B, C or AA.")
            return

        confirmed = messagebox.askyesno(
            "XLSX editor",
            f"Insert column {self.edit_insert_column.get().upper()} into selected XLSX file(s)?",
        )
        if not confirmed:
            return

        if not self.ensure_xlsx_editor_ready(prompt=True):
            return

        try:
            files = self.collect_edit_operation_files()
        except (XlsxEditDependencyError, ValueError) as exc:
            messagebox.showerror("XLSX editor", str(exc))
            return

        self.start_xlsx_edit_operation("insert", files, column=self.edit_insert_column.get())

    def delete_xlsx_column(self) -> None:
        if self.edit_running:
            return
        if not self.edit_source.get():
            messagebox.showerror("XLSX editor", "Choose XLSX source first.")
            return
        if not self.edit_delete_column.get():
            messagebox.showerror("XLSX editor", "Enter delete column, for example A, B, C or AA.")
            return

        if not self.ensure_xlsx_editor_ready(prompt=True):
            return

        try:
            files = self.collect_edit_operation_files()
        except (XlsxEditDependencyError, ValueError) as exc:
            messagebox.showerror("XLSX editor", str(exc))
            return

        confirmed = messagebox.askyesno(
            "XLSX editor",
            f"Delete column {self.edit_delete_column.get().upper()} from {len(files)} selected XLSX file(s)?",
        )
        if not confirmed:
            return

        self.start_xlsx_edit_operation("delete", files, column=self.edit_delete_column.get())

    def swap_xlsx_columns(self) -> None:
        if self.edit_running:
            return
        if not self.edit_source.get():
            messagebox.showerror("XLSX editor", "Choose XLSX source first.")
            return
        if not self.edit_swap_first_column.get() or not self.edit_swap_second_column.get():
            messagebox.showerror("XLSX editor", "Enter both columns to swap, for example A and D.")
            return

        if not self.ensure_xlsx_editor_ready(prompt=True):
            return

        try:
            files = self.collect_edit_operation_files()
        except (XlsxEditDependencyError, ValueError) as exc:
            messagebox.showerror("XLSX editor", str(exc))
            return

        first_column = self.edit_swap_first_column.get().upper()
        second_column = self.edit_swap_second_column.get().upper()
        confirmed = messagebox.askyesno(
            "XLSX editor",
            f"Swap columns {first_column} and {second_column} in {len(files)} selected XLSX file(s)?",
        )
        if not confirmed:
            return

        self.start_xlsx_edit_operation(
            "swap",
            files,
            column=f"{self.edit_swap_first_column.get()}:{self.edit_swap_second_column.get()}",
        )

    def start_xlsx_edit_operation(
        self,
        operation: str,
        files: list[Path],
        *,
        column: str = "",
        show_done: bool = True,
    ) -> None:
        self.cancel_edit_event.clear()
        self.drain_edit_queue()
        self.edit_running = True
        self.set_edit_buttons_state(tk.DISABLED)
        self.edit_cancel_button.configure(state=tk.NORMAL)
        self.edit_progress.configure(value=0)
        self.edit_progress_text.set("Progress: 0%")
        self.edit_current_file.set("Current file: preparing...")
        self.edit_status.set(f"{operation.title()} started.")

        thread = threading.Thread(
            target=self.run_xlsx_edit_operation,
            args=(operation, files, self.edit_sheet.get() or None, column, self.edit_header.get(), show_done),
            daemon=True,
        )
        thread.start()
        self.after(100, self.process_edit_queue)

    def run_xlsx_edit_operation(
        self,
        operation: str,
        files: list[Path],
        sheet: str | None,
        column: str,
        header: str,
        show_done: bool,
    ) -> None:
        reports = []
        results = []
        total = len(files)
        try:
            backend = self.load_xlsx_edit_backend()
            for index, file_path in enumerate(files, start=1):
                if self.cancel_edit_event.is_set():
                    self.edit_queue.put(("cancelled", len(results) or len(reports)))
                    return

                self.edit_queue.put(("progress", (index, total, file_path, operation, 0)))
                if operation == "check":
                    reports.append(backend.inspect_xlsx_columns(file_path, sheet=sheet))
                elif operation == "insert":
                    results.append(backend.insert_column_in_xlsx(file_path, column, sheet=sheet, header=header))
                elif operation == "delete":
                    results.append(backend.delete_column_in_xlsx(file_path, column, sheet=sheet))
                elif operation == "swap":
                    first_column, second_column = column.split(":", maxsplit=1)
                    results.append(backend.swap_columns_in_xlsx(file_path, first_column, second_column, sheet=sheet))
                else:
                    raise ValueError(f"Unknown operation: {operation}")
                self.edit_queue.put(("progress", (index, total, file_path, operation, 100)))
        except (XlsxEditDependencyError, ValueError) as exc:
            self.edit_queue.put(("error", str(exc)))
            return
        except OSError as exc:
            self.edit_queue.put(("error", str(exc)))
            return

        if operation == "check":
            self.edit_queue.put(("done_check", (reports, show_done)))
        else:
            self.edit_queue.put(("done_edit", (operation, len(results))))

    def process_edit_queue(self) -> None:
        while True:
            try:
                event_type, payload = self.edit_queue.get_nowait()
            except queue.Empty:
                break

            if event_type == "progress":
                index, total, file_path, operation, percent = payload
                self.edit_progress.configure(value=percent)
                self.edit_progress_text.set(f"Progress: {percent}%")
                self.edit_current_file.set(f"Current file: {index}/{total} - {file_path.name}")
                self.edit_status.set(f"{operation.title()}: {file_path.name}")
            elif event_type == "error":
                self.finish_xlsx_edit_operation()
                self.edit_status.set("XLSX editor operation failed.")
                messagebox.showerror("XLSX editor", str(payload))
                return
            elif event_type == "cancelled":
                self.finish_xlsx_edit_operation()
                self.edit_status.set("XLSX editor operation cancelled.")
                self.edit_current_file.set("Current file: cancelled.")
                messagebox.showinfo("XLSX editor", f"Operation cancelled.\nProcessed: {payload} file(s).")
                return
            elif event_type == "done_check":
                reports, show_done = payload
                self.populate_edit_table(reports)
                self.finish_xlsx_edit_operation()
                self.edit_status.set(f"Checked: {len(reports)} file(s).")
                if show_done:
                    messagebox.showinfo("XLSX editor", f"Column check completed.\nChecked: {len(reports)} file(s).")
                return
            elif event_type == "done_edit":
                operation, count = payload
                self.finish_xlsx_edit_operation()
                verbs = {"insert": "Inserted", "delete": "Deleted", "swap": "Swapped"}
                verb = verbs.get(operation, operation.title())
                self.edit_status.set(f"{verb} column in {count} file(s).")
                messagebox.showinfo("XLSX editor", f"Completed.\nUpdated: {count} file(s).")
                return
        if self.edit_running:
            self.after(100, self.process_edit_queue)

    def finish_xlsx_edit_operation(self) -> None:
        self.edit_running = False
        self.set_edit_buttons_state(tk.NORMAL)
        self.edit_cancel_button.configure(state=tk.DISABLED)
        self.edit_progress.configure(value=100)
        self.edit_progress_text.set("Progress: 100%")

    def set_edit_buttons_state(self, state: str) -> None:
        self.edit_check_button.configure(state=state)
        self.edit_insert_button.configure(state=state)
        self.edit_delete_button.configure(state=state)
        self.edit_swap_button.configure(state=state)

    def cancel_xlsx_edit(self) -> None:
        if not self.edit_running:
            return
        self.cancel_edit_event.set()
        self.edit_cancel_button.configure(state=tk.DISABLED)
        self.edit_status.set("Cancel requested. Waiting for current file to finish...")

    def drain_edit_queue(self) -> None:
        while True:
            try:
                self.edit_queue.get_nowait()
            except queue.Empty:
                return

    def collect_edit_source_files(self) -> list[Path]:
        source = Path(self.edit_source.get())
        if source.is_file():
            return [source]
        if source.is_dir():
            pattern = "**/*.xlsx" if self.edit_recursive.get() else "*.xlsx"
            files = sorted(path for path in source.glob(pattern) if path.is_file() and not path.name.startswith("~$"))
            if files:
                return files
            raise ValueError(f"No .xlsx files found in: {source}")
        raise ValueError(f"Source must be an XLSX file or directory: {source}")

    def collect_edit_operation_files(self) -> list[Path]:
        if self.edit_checked_sources:
            return [Path(source) for source in sorted(self.edit_checked_sources)]
        if self.edit_table.get_children():
            raise ValueError("Select at least one row in the table, or clear/check columns again.")
        return self.collect_edit_source_files()

    def toggle_edit_row(self, event: object) -> None:
        if self.edit_running:
            return
        row_id = self.edit_table.identify_row(event.y)
        if not row_id:
            return
        values = list(self.edit_table.item(row_id, "values"))
        if row_id in self.edit_checked_sources:
            self.edit_checked_sources.remove(row_id)
            values[0] = "[ ]"
        else:
            self.edit_checked_sources.add(row_id)
            values[0] = "[x]"
        self.edit_table.item(row_id, values=values)
        self.sync_edit_select_all_checkbox()

    def toggle_all_edit_rows(self) -> None:
        if self.edit_select_all.get():
            self.select_all_edit_rows()
        else:
            self.clear_edit_rows()

    def select_all_edit_rows(self) -> None:
        for row_id in self.edit_table.get_children():
            self.edit_checked_sources.add(row_id)
            values = list(self.edit_table.item(row_id, "values"))
            values[0] = "[x]"
            self.edit_table.item(row_id, values=values)
        self.edit_select_all.set(bool(self.edit_table.get_children()))

    def clear_edit_rows(self) -> None:
        self.edit_checked_sources.clear()
        for row_id in self.edit_table.get_children():
            values = list(self.edit_table.item(row_id, "values"))
            values[0] = "[ ]"
            self.edit_table.item(row_id, values=values)
        self.edit_select_all.set(False)

    def sync_edit_select_all_checkbox(self) -> None:
        rows = self.edit_table.get_children()
        self.edit_select_all.set(bool(rows) and len(self.edit_checked_sources) == len(rows))

    def check_xlsx_editor_dependency_on_startup(self) -> None:
        if self.xlsx_edit_backend is not None:
            return
        try:
            self.load_xlsx_edit_backend()
            self.edit_status.set("XLSX editor is ready.")
        except XlsxEditDependencyError:
            self.edit_status.set("XLSX editor needs openpyxl. Install requirements manually.")

    def ensure_xlsx_editor_ready(self, *, prompt: bool) -> bool:
        if self.xlsx_edit_backend is not None:
            return True

        try:
            self.load_xlsx_edit_backend()
            self.edit_status.set("XLSX editor is ready.")
            return True
        except XlsxEditDependencyError as exc:
            if not prompt:
                self.edit_status.set(str(exc))
                return False
            messagebox.showerror(
                "XLSX editor",
                "Для работы XLSX editor нужна библиотека openpyxl.\n\n"
                "Установите зависимости вручную из папки приложения:\n"
                "python -m pip install -r requirements.txt",
            )
            self.edit_status.set("XLSX editor is waiting for manually installed requirements.")
            return False

    def load_xlsx_edit_backend(self):
        if self.xlsx_edit_backend is not None:
            return self.xlsx_edit_backend
        try:
            from . import xlsx_edit
        except ModuleNotFoundError as exc:
            if exc.name == "openpyxl":
                raise XlsxEditDependencyError(
                    "XLSX editor requires openpyxl. Install it with: python -m pip install -r requirements.txt"
                ) from exc
            raise
        self.xlsx_edit_backend = xlsx_edit
        return self.xlsx_edit_backend


    def choose_csv_edit_source_file(self) -> None:
        filename = filedialog.askopenfilename(title="Choose CSV file", filetypes=[("CSV files", "*.csv"), ("All files", "*.*")])
        if filename:
            self.csv_edit_source.set(filename)

    def choose_csv_edit_source_dir(self) -> None:
        directory = filedialog.askdirectory(title="Choose folder with CSV files")
        if directory:
            self.csv_edit_source.set(directory)

    def check_csv_columns(self) -> None:
        if self.csv_edit_running:
            return
        if not self.csv_edit_source.get():
            messagebox.showerror("CSV editor", "Choose CSV source first.")
            return
        try:
            files = self.collect_csv_edit_source_files()
        except ValueError as exc:
            messagebox.showerror("CSV editor", str(exc))
            return
        for item in self.csv_edit_table.get_children():
            self.csv_edit_table.delete(item)
        self.csv_checked_sources.clear()
        self.csv_edit_select_all.set(False)
        self.start_csv_edit_operation("check", files)

    def insert_csv_column(self) -> None:
        self.run_csv_edit_action("insert", self.csv_edit_insert_column.get())

    def delete_csv_column(self) -> None:
        self.run_csv_edit_action("delete", self.csv_edit_delete_column.get())

    def swap_csv_columns(self) -> None:
        first = self.csv_edit_swap_first_column.get()
        second = self.csv_edit_swap_second_column.get()
        if not first or not second:
            messagebox.showerror("CSV editor", "Enter both columns to swap, for example A and D.")
            return
        self.run_csv_edit_action("swap", f"{first}:{second}")

    def run_csv_edit_action(self, operation: str, column: str) -> None:
        if self.csv_edit_running:
            return
        if not self.csv_edit_source.get():
            messagebox.showerror("CSV editor", "Choose CSV source first.")
            return
        if not column:
            messagebox.showerror("CSV editor", "Enter a column address, for example A, B, C or AA.")
            return
        try:
            files = self.collect_csv_edit_operation_files()
        except ValueError as exc:
            messagebox.showerror("CSV editor", str(exc))
            return
        if not messagebox.askyesno("CSV editor", f"Apply {operation} to {len(files)} selected CSV file(s)?"):
            return
        self.start_csv_edit_operation(operation, files, column=column)

    def start_csv_edit_operation(self, operation: str, files: list[Path], *, column: str = "") -> None:
        delimiter = "\t" if self.csv_edit_delimiter.get() == "\\t" else self.csv_edit_delimiter.get()
        self.cancel_csv_edit_event.clear()
        self.drain_csv_edit_queue()
        self.csv_edit_running = True
        self.set_csv_edit_buttons_state(tk.DISABLED)
        self.csv_edit_cancel_button.configure(state=tk.NORMAL)
        self.csv_edit_progress.configure(value=0)
        self.csv_edit_progress_text.set("Progress: 0%")
        self.csv_edit_current_file.set("Current file: preparing...")
        self.csv_edit_status.set(f"{operation.title()} started.")
        threading.Thread(
            target=self.run_csv_edit_operation,
            args=(operation, files, column, self.csv_edit_header.get(), delimiter, self.csv_edit_encoding.get()),
            daemon=True,
        ).start()
        self.after(100, self.process_csv_edit_queue)

    def run_csv_edit_operation(self, operation: str, files: list[Path], column: str, header: str, delimiter: str, encoding: str) -> None:
        reports, results = [], []
        total = len(files)
        try:
            from . import csv_edit
            for index, file_path in enumerate(files, start=1):
                if self.cancel_csv_edit_event.is_set():
                    self.csv_edit_queue.put(("cancelled", len(results) or len(reports)))
                    return
                self.csv_edit_queue.put(("progress", (index, total, file_path, operation, 0)))
                if operation == "check":
                    reports.append(csv_edit.inspect_csv_columns(file_path, delimiter=delimiter, encoding=encoding))
                elif operation == "insert":
                    results.append(csv_edit.insert_column_in_csv(file_path, column, header=header, delimiter=delimiter, encoding=encoding))
                elif operation == "delete":
                    results.append(csv_edit.delete_column_in_csv(file_path, column, delimiter=delimiter, encoding=encoding))
                elif operation == "swap":
                    first, second = column.split(":", maxsplit=1)
                    results.append(csv_edit.swap_columns_in_csv(file_path, first, second, delimiter=delimiter, encoding=encoding))
                else:
                    raise ValueError(f"Unknown operation: {operation}")
                self.csv_edit_queue.put(("progress", (index, total, file_path, operation, 100)))
        except Exception as exc:
            self.csv_edit_queue.put(("error", str(exc)))
            return
        self.csv_edit_queue.put(("done_check" if operation == "check" else "done_edit", reports if operation == "check" else (operation, len(results))))

    def process_csv_edit_queue(self) -> None:
        while True:
            try:
                event_type, payload = self.csv_edit_queue.get_nowait()
            except queue.Empty:
                break
            if event_type == "progress":
                index, total, file_path, operation, percent = payload
                self.csv_edit_progress.configure(value=percent)
                self.csv_edit_progress_text.set(f"Progress: {percent}%")
                self.csv_edit_current_file.set(f"Current file: {index}/{total} - {file_path.name}")
                self.csv_edit_status.set(f"{operation.title()}: {file_path.name}")
            elif event_type == "error":
                self.finish_csv_edit_operation()
                self.csv_edit_status.set("CSV editor operation failed.")
                messagebox.showerror("CSV editor", str(payload))
                return
            elif event_type == "cancelled":
                self.finish_csv_edit_operation()
                self.csv_edit_status.set("CSV editor operation cancelled.")
                messagebox.showinfo("CSV editor", f"Operation cancelled.\nProcessed: {payload} file(s).")
                return
            elif event_type == "done_check":
                self.populate_csv_edit_table(payload)
                self.finish_csv_edit_operation()
                self.csv_edit_status.set(f"Checked: {len(payload)} file(s).")
                messagebox.showinfo("CSV editor", f"Column check completed.\nChecked: {len(payload)} file(s).")
                return
            elif event_type == "done_edit":
                _operation, count = payload
                self.finish_csv_edit_operation()
                self.csv_edit_status.set(f"Completed: {count} file(s).")
                messagebox.showinfo("CSV editor", f"Completed.\nUpdated: {count} file(s).")
                return
        if self.csv_edit_running:
            self.after(100, self.process_csv_edit_queue)

    def populate_csv_edit_table(self, reports: list[object]) -> None:
        for report in reports:
            source_key = str(report.source)
            self.csv_checked_sources.add(source_key)
            columns_text = ", ".join(f"{column.letter}: {column.header or '-'}" for column in report.columns)
            self.csv_edit_table.insert("", tk.END, iid=source_key, values=("[x]", report.source.name, report.column_count, columns_text))
        self.sync_csv_edit_select_all_checkbox()

    def finish_csv_edit_operation(self) -> None:
        self.csv_edit_running = False
        self.set_csv_edit_buttons_state(tk.NORMAL)
        self.csv_edit_cancel_button.configure(state=tk.DISABLED)
        self.csv_edit_progress.configure(value=100)
        self.csv_edit_progress_text.set("Progress: 100%")

    def set_csv_edit_buttons_state(self, state: str) -> None:
        self.csv_edit_check_button.configure(state=state)
        self.csv_edit_insert_button.configure(state=state)
        self.csv_edit_delete_button.configure(state=state)
        self.csv_edit_swap_button.configure(state=state)

    def cancel_csv_edit(self) -> None:
        if self.csv_edit_running:
            self.cancel_csv_edit_event.set()
            self.csv_edit_cancel_button.configure(state=tk.DISABLED)
            self.csv_edit_status.set("Cancel requested. Waiting for current file to finish...")

    def drain_csv_edit_queue(self) -> None:
        while True:
            try:
                self.csv_edit_queue.get_nowait()
            except queue.Empty:
                return

    def collect_csv_edit_source_files(self) -> list[Path]:
        source = Path(self.csv_edit_source.get())
        if source.is_file():
            return [source]
        if source.is_dir():
            pattern = "**/*.csv" if self.csv_edit_recursive.get() else "*.csv"
            files = sorted(path for path in source.glob(pattern) if path.is_file())
            if files:
                return files
            raise ValueError(f"No .csv files found in: {source}")
        raise ValueError(f"Source must be a CSV file or directory: {source}")

    def collect_csv_edit_operation_files(self) -> list[Path]:
        if self.csv_checked_sources:
            return [Path(source) for source in sorted(self.csv_checked_sources)]
        if self.csv_edit_table.get_children():
            raise ValueError("Select at least one row in the table, or check columns again.")
        return self.collect_csv_edit_source_files()

    def toggle_csv_edit_row(self, event: object) -> None:
        if self.csv_edit_running:
            return
        row_id = self.csv_edit_table.identify_row(event.y)
        if not row_id:
            return
        values = list(self.csv_edit_table.item(row_id, "values"))
        if row_id in self.csv_checked_sources:
            self.csv_checked_sources.remove(row_id)
            values[0] = "[ ]"
        else:
            self.csv_checked_sources.add(row_id)
            values[0] = "[x]"
        self.csv_edit_table.item(row_id, values=values)
        self.sync_csv_edit_select_all_checkbox()

    def toggle_all_csv_edit_rows(self) -> None:
        rows = self.csv_edit_table.get_children()
        if self.csv_edit_select_all.get():
            self.csv_checked_sources.update(rows)
            marker = "[x]"
        else:
            self.csv_checked_sources.clear()
            marker = "[ ]"
        for row_id in rows:
            values = list(self.csv_edit_table.item(row_id, "values"))
            values[0] = marker
            self.csv_edit_table.item(row_id, values=values)

    def sync_csv_edit_select_all_checkbox(self) -> None:
        rows = self.csv_edit_table.get_children()
        self.csv_edit_select_all.set(bool(rows) and len(self.csv_checked_sources) == len(rows))


def main() -> int:
    app = FileToolsApp()
    app.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

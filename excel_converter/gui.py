"""Tk desktop UI. All file operations run locally."""

from __future__ import annotations

import copy
import json
import queue
import threading
import tempfile
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from openpyxl.utils import get_column_letter

from .core import (convert, export_result, headers, load_profile, load_source,
                   suggest_header, suggest_mapping, write_report)
from .demo import create_demo, create_scenario
from .dpi import create_root, px, scale_widgets, apply_style_scale, watch_dpi


class Application(ttk.Frame):
    def __init__(self, root):
        super().__init__(root, padding=18)
        self.root = root
        self.pack(fill="both", expand=True)
        self.source = None
        self.result = None
        self.corrections = {}
        self.profile = load_profile()
        self.mapping_vars = {}
        self.mapping_options = []
        self.busy = False
        self.events = queue.Queue()
        self.demo_key = None
        self.demo_directories = []
        self.sheet_var = tk.StringVar()
        self.header_var = tk.StringVar(value="1")
        self.status = tk.StringVar(value="Откройте файл или создайте учебные примеры.")
        self.filename = tk.StringVar(value="Файл не выбран")
        self.profile_label = tk.StringVar()
        self.build_ui()
        scale_widgets(self)
        apply_style_scale(root)
        self.refresh_profile_label()
        self.header_var.trace_add("write", lambda *_: self.invalidate(clear_mapping=True))
        self.root.after(100, self.poll)

    def build_ui(self):
        from .layout import build
        build(self)

    @staticmethod
    def make_tree(parent):
        frame = ttk.Frame(parent)
        frame.pack(fill="both", expand=True)
        tree = ttk.Treeview(frame, show="headings")
        vertical = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        horizontal = ttk.Scrollbar(frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vertical.set, xscrollcommand=horizontal.set)
        tree.grid(row=0, column=0, sticky="nsew")
        vertical.grid(row=0, column=1, sticky="ns")
        horizontal.grid(row=1, column=0, sticky="ew")
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)
        return tree

    @staticmethod
    def fill_tree(tree, columns, rows):
        tree.delete(*tree.get_children())
        ids = [str(i) for i in range(len(columns))]
        tree.configure(columns=ids)
        for key, title in zip(ids, columns):
            tree.heading(key, text=title)
            tree.column(key, width=px(tree, 175 if key != "0" else 95), minwidth=px(tree, 65), stretch=False)
        tree.tag_configure('even', background='#f7f9fc')
        for index, row in enumerate(rows):
            tree.insert("", "end", values=["" if v is None else str(v) for v in row], tags=('even',) if index % 2 else ())

    def refresh_profile_label(self):
        self.profile_label.set(self.profile["name"] + " · " + self.profile.get("description", "Пользовательский профиль"))
        dates = next((f for f in self.profile['fields'] if f['kind'] == 'date'), {})
        amounts = next((f for f in self.profile['fields'] if f['kind'] == 'amount'), {})
        self.date_order.set({'auto': 'Авто', 'dmy': 'День–месяц–год', 'mdy': 'Месяц–день–год'}[dates.get('date_order', 'auto')])
        self.decimal_separator.set({'auto': 'Авто', '.': 'Точка', ',': 'Запятая'}[amounts.get('decimal_separator', 'auto')])
        self.drop_time.set(dates.get('drop_time', False))

    def parsing_changed(self):
        for field in self.profile['fields']:
            if field['kind'] == 'date':
                field['date_order'] = {'Авто': 'auto', 'День–месяц–год': 'dmy', 'Месяц–день–год': 'mdy'}[self.date_order.get()]
                field['drop_time'] = self.drop_time.get()
            elif field['kind'] == 'amount':
                field['decimal_separator'] = {'Авто': 'auto', 'Точка': '.', 'Запятая': ','}[self.decimal_separator.get()]
        self.invalidate(keep_corrections=True)
        self.status.set('Правила распознавания изменены. Выполните проверку заново.')

    def invalidate(self, clear_mapping=False, keep_corrections=False):
        if not keep_corrections:
            self.corrections = {}
        self.result = None
        self.fix_button.configure(state="disabled")
        self.export_button.configure(state="disabled")
        self.report_button.configure(state="disabled")
        self.result_tree.delete(*self.result_tree.get_children())
        self.issues_tree.delete(*self.issues_tree.get_children())
        self.changes_tree.delete(*self.changes_tree.get_children())
        self.demo_fix_button.configure(state='disabled')
        for variable in self.metric_vars:
            variable.set('—')
        self.check_hint.set('Настройки изменились. Проверьте данные заново.')
        self.summary.set("Настройки изменены. Выполните проверку заново.")
        if clear_mapping:
            self.mapping_vars.clear()
            for widget in self.mapping_frame.winfo_children():
                widget.destroy()

    def set_busy(self, busy):
        self.busy = busy
        for button in (self.open_button, self.other_file_button, self.demo_button, self.profile_button, *self.demo_buttons):
            button.configure(state="disabled" if busy else "normal")
        for button in (self.save_profile_button, self.mapping_button, self.check_button, self.next_button):
            button.configure(state='normal' if not busy and self.source else 'disabled')
        self.demo_fix_button.configure(state='normal' if not busy and self.demo_key == 'errors' and self.result and self.result.errors else 'disabled')
        if self.demo_key == 'errors':
            self.demo_fix_button.pack(side='left', padx=8)
        else:
            self.demo_fix_button.pack_forget()
        if busy:
            self.progress.start(12)
        else:
            self.progress.stop()
        self.sheet_box.configure(state="disabled" if busy else "readonly")
        self.header_entry.configure(state="disabled" if busy else "normal")
        self.date_order_box.configure(state='disabled' if busy else 'readonly')
        self.decimal_box.configure(state='disabled' if busy else 'readonly')
        self.time_check.configure(state='disabled' if busy else 'normal')
        for widget in self.mapping_frame.winfo_children():
            if isinstance(widget, ttk.Combobox):
                widget.configure(state="disabled" if busy else "readonly")
        self.export_button.configure(state="normal" if not busy and self.result and not self.result.errors else "disabled")
        self.report_button.configure(state="normal" if not busy and self.result else "disabled")
        self.fix_button.configure(state="normal" if not busy and self.result and any(i.field_id for i in self.result.errors) else "disabled")

    def run_background(self, operation, on_success, text):
        if self.busy:
            return
        self.set_busy(True)
        self.status.set(text)
        def work():
            try:
                self.events.put((True, operation(), on_success))
            except Exception as exc:
                self.events.put((False, str(exc), on_success))
        threading.Thread(target=work, daemon=True).start()

    def poll(self):
        try:
            success, value, callback = self.events.get_nowait()
            self.set_busy(False)
            if success:
                callback(value)
            else:
                self.status.set("Операция не выполнена.")
                messagebox.showerror("Ошибка", value)
        except queue.Empty:
            pass
        self.root.after(100, self.poll)

    def open_file(self):
        if self.busy:
            return
        path = filedialog.askopenfilename(filetypes=[("Excel / CSV", "*.xlsx *.csv"), ("Все файлы", "*.*")])
        if path:
            self.load_file(path)

    def load_file(self, path):
        self.run_background(lambda: load_source(path), self.loaded, "Чтение файла…")

    def loaded(self, source):
        self.demo_key = None
        self.source = source
        self.invalidate(clear_mapping=True)
        self.filename.set(source.path.name)
        self.sheet_box.configure(values=list(source.sheets))
        self.sheet_var.set(next(iter(source.sheets)))
        self.select_sheet()
        self.notebook.select(self.source_tab)
        self.set_busy(False)

    def select_sheet(self):
        if not self.source:
            return
        rows = self.source.sheets[self.sheet_var.get()]
        suggestion = suggest_header(rows, self.profile)
        self.header_var.set(str(suggestion or 1))
        width = max((len(row) for row in rows[:200]), default=0)
        self.fill_tree(self.source_tree, ["Строка"] + [get_column_letter(i + 1) for i in range(width)],
                       [[i, *row] for i, row in enumerate(rows[:200], 1)])
        if rows:
            self.apply_headers()
        self.status.set(f"Лист «{self.sheet_var.get()}»: {len(rows)} строк. " +
                        (f"Предложена строка заголовков {suggestion}; проверьте её." if suggestion else "Заголовок не определён однозначно. Укажите строку вручную."))

    def apply_headers(self):
        if not self.source:
            return
        try:
            names = headers(self.source.sheets[self.sheet_var.get()], int(self.header_var.get()))
        except (ValueError, KeyError) as exc:
            messagebox.showerror("Заголовки", str(exc))
            return
        self.invalidate(clear_mapping=True)
        self.mapping_options = ["— не выбрано —"] + [f"{get_column_letter(i + 1)} · {name or '(без заголовка)'}" for i, name in enumerate(names)]
        suggestion = suggest_mapping(names, self.profile)
        for row, spec in enumerate(self.profile["fields"]):
            ttk.Label(self.mapping_frame, text=spec["title"] + (" *" if spec["required"] else ""), width=27).grid(row=row, column=0, sticky="w", pady=7)
            index = suggestion[spec["id"]]
            variable = tk.StringVar(value=self.mapping_options[index + 1 if index is not None else 0])
            self.mapping_vars[spec["id"]] = variable
            box = ttk.Combobox(self.mapping_frame, values=self.mapping_options, textvariable=variable, state="readonly", width=56)
            box.grid(row=row, column=1, sticky="w", padx=12)
            box.bind("<<ComboboxSelected>>", lambda _: self.invalidate())
        scale_widgets(self.mapping_frame)
        self.status.set("Сопоставление предложено. Проверьте его перед преобразованием.")

    def get_mapping(self):
        if not self.source or not self.mapping_vars:
            raise ValueError("Откройте файл и примените строку заголовков.")
        return {key: (self.mapping_options.index(var.get()) - 1 if self.mapping_options.index(var.get()) else None)
                for key, var in self.mapping_vars.items()}

    def check(self):
        if self.busy:
            return
        try:
            mapping = self.get_mapping()
            header = int(self.header_var.get())
        except ValueError as exc:
            messagebox.showerror("Проверка", str(exc))
            return
        self.invalidate(keep_corrections=True)
        source, sheet, profile = self.source, self.sheet_var.get(), copy.deepcopy(self.profile)
        corrections = dict(self.corrections)
        self.run_background(lambda: convert(source, sheet, header, profile, mapping, corrections), self.checked, "Проверка и преобразование…")

    def checked(self, result):
        self.result = result
        self.corrections = dict(result.corrections)
        warnings = len(result.issues) - len(result.errors)
        self.summary.set(f"Прочитано записей: {result.read_rows} • Корректных: {len(result.records)} • Ошибочных строк: {len(result.invalid_rows)} • Пустых: {len(result.blank_rows)}\n"
                         f"Ошибок: {len(result.errors)} • Предупреждений: {warnings} • Преобразований и правок: {len(result.changes)}. Предпросмотр первых 200 корректных записей.")
        self.fill_tree(self.result_tree, ["Исх. строка"] + [f["title"] for f in result.profile["fields"]],
                       [[number, *record.values()] for number, record in list(zip(result.source_rows, result.records))[:200]])
        self.fill_tree(self.issues_tree, ["Уровень", "Строка", "Поле", "Описание"],
                       [["Ошибка" if issue.severity == "error" else "Внимание", issue.row, issue.field, issue.message] for issue in result.issues[:2000]])
        self.issues_tree.column("3", width=px(self, 740))
        for item, issue in zip(self.issues_tree.get_children(), result.issues):
            self.issues_tree.item(item, tags=(issue.severity,))
        for variable, value in zip(self.metric_vars, [result.read_rows, len(result.records), len(result.errors), len(result.changes)]):
            variable.set(str(value))
        titles = {f['id']: f['title'] for f in result.profile['fields']}
        self.fill_tree(self.changes_tree, ['Строка', 'Поле', 'Было', 'Стало', 'Способ'],
                       [[c['row'], titles.get(c['field'], c['field']), c['before'], c['after'], 'Вручную' if c['kind'] == 'manual' else 'Автоматически'] for c in result.changes[:2000]])
        self.changes_tree.column('2', width=px(self, 260))
        self.changes_tree.column('3', width=px(self, 260))
        self.check_hint.set(f'Ошибок: {len(result.errors)}. Предупреждений: {warnings}. ' +
                            ('Исправьте красные строки. Жёлтые строки требуют просмотра.' if result.errors else 'Ошибок нет — результат можно сохранить.') +
                            (' Учебный пример: можно применить заранее подготовленные исправления.' if self.demo_key == 'errors' and result.errors else ''))
        self.set_busy(False)
        self.status.set(("Выгрузка заблокирована. Нажмите «Исправить ошибки» на вкладке «Проверки»; ошибки сопоставления исправляются в настройках." if result.errors else "Проверка завершена. Результат готов к сохранению.") +
                        (" Показаны первые 2000 замечаний; полный список — в отчёте." if len(result.issues) > 2000 else ""))
        self.notebook.select(self.issues_tab if result.errors else self.result_tab)

    def show_issue(self, _event):
        selection = self.issues_tree.selection()
        if selection:
            if self.result and not self.busy:
                issue = self.result.issues[self.issues_tree.index(selection[0])]
                if issue.severity == 'error' and issue.field_id:
                    dialog = self.edit_errors()
                    for item, candidate in zip(dialog.tree.get_children(), dialog.items):
                        if candidate.row == issue.row and candidate.field_id == issue.field_id:
                            dialog.tree.selection_set(item)
                            dialog.tree.see(item)
                            dialog.select()
                            break
                    return
            messagebox.showinfo("Замечание", "\n".join(str(v) for v in self.issues_tree.item(selection[0], "values")))

    def edit_errors(self):
        if self.busy or not self.result or not any(i.field_id for i in self.result.errors):
            return
        from .corrections import CorrectionDialog
        return CorrectionDialog(self)

    def export(self):
        if not self.result or self.result.errors:
            return
        warnings = [i for i in self.result.issues if i.severity == "warning"]
        if warnings and not messagebox.askyesno("Проверьте предупреждения", f"Есть предупреждения: {len(warnings)}. Они доступны на вкладке «Проверки».\n\nВы просмотрели их и хотите сохранить результат?"):
            self.notebook.select(self.issues_tab)
            return
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile=self.source.path.stem + "_converted.xlsx", filetypes=[("Excel", "*.xlsx")])
        if path:
            result = self.result
            self.run_background(lambda: export_result(result, path), self.exported, "Сохранение результата и отчёта…")

    def exported(self, paths):
        self.status.set(f"Сохранено: {paths[0]}")
        messagebox.showinfo("Готово", f"Результат:\n{paths[0]}\n\nОтчёт:\n{paths[1]}")

    def save_report(self):
        if not self.result:
            return
        path = filedialog.asksaveasfilename(defaultextension=".json", initialfile="validation.report.json", filetypes=[("JSON", "*.json")])
        if path:
            try:
                write_report(self.result, Path(path))
                self.status.set(f"Отчёт сохранён: {path}")
            except Exception as exc:
                messagebox.showerror("Отчёт", str(exc))

    def open_profile(self):
        path = filedialog.askopenfilename(filetypes=[("JSON-профиль", "*.json")])
        if not path:
            return
        try:
            profile = load_profile(path)
        except Exception as exc:
            messagebox.showerror("Профиль", str(exc))
            return
        self.profile = profile
        self.refresh_profile_label()
        self.invalidate(clear_mapping=True)
        if self.source:
            self.select_sheet()

    def save_profile(self):
        try:
            mapping = self.get_mapping()
            names = headers(self.source.sheets[self.sheet_var.get()], int(self.header_var.get()))
            if any(index is not None and (not names[index] or names.count(names[index]) != 1) for index in mapping.values()):
                raise ValueError("Для сохранения профиля выбранные заголовки должны быть непустыми и уникальными.")
            profile = copy.deepcopy(self.profile)
            profile["mapping"] = {key: names[index] if index is not None else None for key, index in mapping.items()}
            path = filedialog.asksaveasfilename(defaultextension=".json", initialfile="custom_profile.json", filetypes=[("JSON-профиль", "*.json")])
            if path:
                with Path(path).open("x", encoding="utf-8") as stream:
                    json.dump(profile, stream, ensure_ascii=False, indent=2)
                self.status.set(f"Профиль сохранён: {path}")
        except Exception as exc:
            messagebox.showerror("Профиль", str(exc))

    def make_demo(self):
        directory = filedialog.askdirectory(title="Папка для синтетических примеров")
        if directory:
            self.run_background(lambda: create_demo(directory), lambda paths: self.load_file(paths[0]), "Создание учебных примеров…")

    def toggle_fullscreen(self):
        self.root.attributes('-fullscreen', not self.root.attributes('-fullscreen'))

    def start_demo(self, scenario):
        if self.busy:
            return
        directory = tempfile.TemporaryDirectory(prefix='excel-converter-demo-')
        self.demo_directories.append(directory)
        def prepare():
            return load_source(create_scenario(directory.name, scenario))
        def ready(source):
            self.profile = load_profile()
            self.refresh_profile_label()
            self.loaded(source)
            self.demo_key = scenario
            self.filename.set('УЧЕБНЫЙ ПРИМЕР · ' + source.path.name)
            self.check()
        self.run_background(prepare, ready, 'Подготовка синтетического примера…')

    def fix_demo(self):
        if self.busy or self.demo_key != 'errors' or not self.result:
            return
        # Only synthetic scenarios expose this shortcut; real files use the cell editor.
        self.corrections.update({(2, 'date'): '2026-05-04', (2, 'identifier'): '000000012345', (2, 'amount'): '1234.00'})
        self.check()


def launch():
    root = create_root()
    Application(root)
    watch_dpi(root)
    root.mainloop()

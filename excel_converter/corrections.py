"""Modal editor for erroneous cells; edits remain staged until applied."""
import tkinter as tk
from tkinter import ttk, messagebox

from .core import UnsafeCell, parse_value
from .dpi import fit_window, px, scale_widgets


class CorrectionDialog(tk.Toplevel):
    def __init__(self, app):
        super().__init__(app.root)
        self.app = app
        self.result = app.result
        self.pending = dict(self.result.corrections)
        self.items = [i for i in self.result.errors if i.field_id is not None]
        self.title("Исправление ошибочных ячеек")
        fit_window(self, 1050, 640)
        self.transient(app.root)
        self.grab_set()
        body = ttk.Frame(self, padding=15)
        body.pack(fill="both", expand=True)
        ttk.Label(body, text="Выберите ошибку, введите значение и нажмите «Запомнить». Можно исправить несколько ячеек перед повторной проверкой.", wraplength=990).pack(anchor="w")
        self.tree = app.make_tree(body)
        app.fill_tree(self.tree, ["Строка", "Поле", "Исходное значение", "Исправление", "Ошибка"],
                      [[i.row, i.field, self.original(i), self.pending.get((i.row, i.field_id), ""), i.message] for i in self.items])
        self.tree.column("4", width=px(self, 500))
        self.tree.bind("<<TreeviewSelect>>", self.select)
        self.detail = tk.StringVar()
        ttk.Label(body, textvariable=self.detail, wraplength=990).pack(anchor="w", pady=8)
        self.value = tk.StringVar()
        ttk.Entry(body, textvariable=self.value, width=90).pack(fill="x", pady=5)
        buttons = ttk.Frame(body)
        buttons.pack(fill="x", pady=8)
        ttk.Button(buttons, text="Запомнить", command=self.remember).pack(side="left")
        ttk.Button(buttons, text="Для всех таких же ошибок", command=lambda: self.remember(True)).pack(side="left", padx=8)
        ttk.Button(buttons, text="Отменить правку ячейки", command=self.reset_cell).pack(side="left")
        ttk.Button(buttons, text="Применить и проверить", command=self.apply).pack(side="right")
        ttk.Button(buttons, text="Закрыть без применения", command=self.destroy).pack(side="right", padx=8)
        self.note = tk.StringVar(value="Исходный файл не изменяется. Правки попадут в отчёт обработки.")
        ttk.Label(body, textvariable=self.note, wraplength=990).pack(anchor="w")
        children = self.tree.get_children()
        if children:
            self.tree.selection_set(children[0])
            self.select()
        scale_widgets(self)

    def original(self, issue):
        row = self.result.source.sheets[self.result.sheet][issue.row - 1]
        index = self.result.mapping[issue.field_id]
        return row[index] if index is not None and index < len(row) else None

    def selected(self):
        selection = self.tree.selection()
        return self.items[self.tree.index(selection[0])] if selection else None

    def select(self, _event=None):
        issue = self.selected()
        if issue:
            value = self.pending.get((issue.row, issue.field_id), self.original(issue))
            self.value.set("" if value is None else str(value))
            self.detail.set(f"Строка {issue.row}, {issue.field}: {issue.message}")

    def remember(self, all_same=False):
        issue = self.selected()
        if not issue:
            return
        if all_same and isinstance(self.original(issue), UnsafeCell):
            messagebox.showinfo("Проверьте каждую ячейку", "Формулы и ошибки Excel исправляются по одной: одинаковое сообщение не означает одинаковое исходное значение.", parent=self)
            return
        spec = next(f for f in self.result.profile["fields"] if f["id"] == issue.field_id)
        value = self.value.get()
        try:
            parse_value(value, spec)
        except ValueError as exc:
            messagebox.showerror("Значение всё ещё некорректно", str(exc), parent=self)
            return
        targets = [i for i in self.items if i.field_id == issue.field_id and type(self.original(i)) is type(self.original(issue)) and self.original(i) == self.original(issue)] if all_same else [issue]
        for target in targets:
            self.pending[(target.row, target.field_id)] = value
        self.refresh()
        self.note.set(f"Запомнено исправлений: {len(targets)}. Нажмите «Применить и проверить», когда закончите.")

    def refresh(self):
        for item, issue in zip(self.tree.get_children(), self.items):
            self.tree.set(item, "3", self.pending.get((issue.row, issue.field_id), ""))

    def reset_cell(self):
        issue = self.selected()
        if issue:
            self.pending.pop((issue.row, issue.field_id), None)
            self.refresh()
            self.select()

    def apply(self):
        # Uncommitted entry text must not be silently lost.
        issue = self.selected()
        if issue:
            saved = self.pending.get((issue.row, issue.field_id), self.original(issue))
            if self.value.get() != ("" if saved is None else str(saved)):
                messagebox.showinfo("Есть незапомненное значение", "Нажмите «Запомнить» перед применением.", parent=self)
                return
        self.app.corrections = dict(self.pending)
        self.destroy()
        self.app.check()

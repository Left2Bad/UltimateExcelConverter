"""Run explicitly with a working Tk runtime; no dialogs or visible windows."""
import tempfile
import tkinter as tk
from pathlib import Path

from excel_converter.core import convert, load_source
from excel_converter.demo import create_demo
from excel_converter.gui import Application


def main():
    with tempfile.TemporaryDirectory() as temp:
        paths = create_demo(temp)
        root = tk.Tk()
        root.withdraw()
        try:
            app = Application(root)
            app.loaded(load_source(paths[0]))
            assert app.header_var.get() == '3'
            assert len(app.mapping_vars) == 7
            result = convert(app.source, app.sheet_var.get(), int(app.header_var.get()), app.profile, app.get_mapping())
            app.checked(result)
            root.update_idletasks()
            assert len(app.result_tree.get_children()) == 3
            assert str(app.export_button['state']) == 'normal'
            app.header_var.set('4')
            assert app.result is None and not app.mapping_vars
            assert str(app.export_button['state']) == 'disabled'
            app.loaded(load_source(paths[1]))
            result = convert(app.source, app.sheet_var.get(), 3, app.profile, app.get_mapping())
            app.checked(result)
            assert str(app.export_button['state']) == 'disabled'
            assert len(app.issues_tree.get_children()) > 0
            print('GUI smoke: PASS (load, mapping, preview, invalidation, blocked export)')
        finally:
            root.destroy()


if __name__ == '__main__':
    main()

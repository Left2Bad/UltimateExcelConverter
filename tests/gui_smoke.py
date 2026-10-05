"""Run explicitly with a working Tk runtime; no dialogs or visible windows."""
import tempfile
import tkinter as tk
import time
from pathlib import Path

from excel_converter.core import convert, load_source
from excel_converter.demo import create_demo
from excel_converter.gui import Application
from excel_converter.dpi import create_root


def main():
    with tempfile.TemporaryDirectory() as temp:
        paths = create_demo(temp)
        root = create_root()
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
            app.source.sheets['Реестр'].append(app.source.sheets['Реестр'][6])
            result = convert(app.source, app.sheet_var.get(), 3, app.profile, app.get_mapping())
            app.checked(result)
            assert str(app.export_button['state']) == 'disabled'
            assert len(app.issues_tree.get_children()) > 0
            dialog = app.edit_errors()
            assert dialog is not None
            dialog.value.set('2026-05-04')
            dialog.remember()
            assert dialog.pending[(7, 'date')] == '2026-05-04'
            dialog.remember(True)
            assert dialog.pending[(10, 'date')] == '2026-05-04'
            dialog.reset_cell()
            assert (7, 'date') not in dialog.pending
            dialog.destroy()
            assert not app.corrections
            def wait_ready():
                deadline = time.monotonic() + 10
                while time.monotonic() < deadline:
                    root.update()
                    if not app.busy and app.result is not None:
                        return
                    time.sleep(0.01)
                raise AssertionError('Demo did not finish')
            for scenario in ['formats', 'csv', 'errors']:
                app.start_demo(scenario)
                wait_ready()
                assert app.demo_key == scenario
                if scenario == 'errors':
                    assert len(app.result.errors) == 3
                    assert str(app.demo_fix_button['state']) == 'normal'
                    app.fix_demo()
                    wait_ready()
                    assert not app.result.errors
                    assert len(app.result.records) == 3
                else:
                    assert not app.result.errors
                assert len(app.changes_tree.get_children()) == len(app.result.changes)
            app.loaded(load_source(paths[0]))
            assert app.demo_key is None
            assert str(app.demo_fix_button['state']) == 'disabled'
            assert not app.corrections
            print('GUI smoke: PASS (load, mapping, preview, invalidation, blocked export)')
        finally:
            root.destroy()


if __name__ == '__main__':
    main()

"""Native DPI context plus sizing regression checks at common Windows scales."""
import ctypes
import sys
from tkinter import font, ttk

from excel_converter.dpi import create_root, px, scale_widgets
from excel_converter.gui import Application


def main():
    root = create_root()
    root.withdraw()
    try:
        if sys.platform == 'win32':
            get_context = ctypes.windll.user32.GetThreadDpiAwarenessContext
            get_context.restype = ctypes.c_void_p
            compare = ctypes.windll.user32.AreDpiAwarenessContextsEqual
            compare.argtypes = [ctypes.c_void_p, ctypes.c_void_p]
            assert compare(get_context(), ctypes.c_void_p(-4)), 'Expected native PerMonitorV2 DPI awareness'
        for percent in (100, 125, 150, 200):
            root._ui_scale = percent / 100
            root.tk.call('tk', 'scaling', root._ui_scale * 96 / 72)
            app = Application(root)
            style = ttk.Style(root)
            line_height = font.Font(root=root, font=style.lookup('Treeview', 'font')).metrics('linespace')
            assert int(style.lookup('Treeview', 'rowheight')) >= line_height + px(root, 4), (percent, line_height)
            app.fill_tree(app.result_tree, ['Строка', 'Поле'], [[1, 'Тест']])
            assert int(app.result_tree.column('1', 'width')) == px(root, 175)
            before = str(app.home_tab.cget('padding'))
            scale_widgets(app)
            assert str(app.home_tab.cget('padding')) == before, 'Repeated scaling must not compound'
            app.destroy()
        print('DPI smoke: PASS (PerMonitorV2; sizing at 100%, 125%, 150%, 200%)')
    finally:
        root.destroy()


if __name__ == '__main__':
    main()

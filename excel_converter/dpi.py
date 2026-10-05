"""Draw Tk at native monitor DPI instead of letting Windows stretch a bitmap."""
import ctypes
import sys
import tkinter as tk
from tkinter import font, ttk


def enable_dpi_awareness():
    """Called before Tk creates an HWND; the packaged EXE also has a manifest."""
    if sys.platform != 'win32':
        return
    user32 = ctypes.windll.user32
    try:
        set_awareness = user32.SetProcessDpiAwarenessContext
        set_awareness.argtypes = [ctypes.c_void_p]
        set_awareness.restype = ctypes.c_bool
        # False also means the EXE manifest already established awareness.
        if set_awareness(ctypes.c_void_p(-4)) or ctypes.GetLastError() == 5:
            return
    except AttributeError:
        pass
    try:
        if ctypes.windll.shcore.SetProcessDpiAwareness(2) in (0, -2147024891):
            return
    except (AttributeError, OSError):
        pass
    user32.SetProcessDPIAware()


def window_dpi(window):
    if sys.platform == 'win32':
        try:
            function = ctypes.windll.user32.GetDpiForWindow
            function.argtypes = [ctypes.c_void_p]
            function.restype = ctypes.c_uint
            dpi = function(window.winfo_id())
            if dpi:
                return dpi
        except AttributeError:
            pass
    return round(float(window.winfo_fpixels('1i')))


def create_root():
    enable_dpi_awareness()
    root = tk.Tk()
    dpi = window_dpi(root)
    root.tk.call('tk', 'scaling', dpi / 72)
    root._ui_scale = dpi / 96
    return root


def px(window, value):
    root = window._root()
    return max(1, round(value * getattr(root, '_ui_scale', 1))) if value else 0


def fit_window(window, width, height, min_width=0, min_height=0):
    # Leave room for the taskbar and decorations on small/high-DPI displays.
    available_width = max(1, window.winfo_screenwidth() - px(window, 32))
    available_height = max(1, window.winfo_screenheight() - px(window, 80))
    window.geometry(f'{min(px(window, width), available_width)}x{min(px(window, height), available_height)}')
    window.minsize(min(px(window, min_width), available_width), min(px(window, min_height), available_height))


def _scaled_value(window, value):
    values = (value,) if isinstance(value, (int, float)) else window.tk.splitlist(value)
    try:
        return tuple(px(window, float(item)) for item in values)
    except (ValueError, TypeError):
        return value


def scale_widgets(window):
    """Scale pixel dimensions, leaving character widths and point fonts alone."""
    if not hasattr(window, '_logical_pixels'):
        options = {}
        for key in ('padding', 'wraplength', 'length'):
            if key in window.keys():
                options[key] = window.cget(key)
        window._logical_pixels = options
        manager = window.winfo_manager()
        info = window.pack_info() if manager == 'pack' else window.grid_info() if manager == 'grid' else {}
        window._logical_spacing = {key: info[key] for key in ('padx', 'pady', 'ipadx', 'ipady') if key in info}
    for key, value in window._logical_pixels.items():
        window.configure(**{key: _scaled_value(window, value)})
    if window._logical_spacing:
        scaled = {key: _scaled_value(window, value) for key, value in window._logical_spacing.items()}
        if window.winfo_manager() == 'pack':
            window.pack_configure(**scaled)
        elif window.winfo_manager() == 'grid':
            window.grid_configure(**scaled)
    for child in window.winfo_children():
        scale_widgets(child)


def apply_style_scale(root):
    style = ttk.Style(root)
    style.configure('TButton', padding=(px(root, 12), px(root, 8)))
    style.configure('Treeview', rowheight=px(root, 29))
    style.configure('Treeview.Heading', padding=px(root, 8))
    style.configure('TNotebook.Tab', padding=(px(root, 16), px(root, 11)))
    style.configure('TScrollbar', arrowsize=px(root, 14))


def watch_dpi(root):
    def update():
        if not root.winfo_exists():
            return
        factor = window_dpi(root) / 96
        old_factor = root._ui_scale
        if abs(factor - old_factor) > .01:
            root._ui_scale = factor
            root.tk.call('tk', 'scaling', factor * 96 / 72)
            # Refresh existing font metrics as well as fonts created later.
            for name in font.names(root):
                item = font.Font(root=root, name=name, exists=True)
                item.configure(size=item.cget('size'))
            def columns(widget):
                if isinstance(widget, ttk.Treeview):
                    for key in widget['columns']:
                        for option in ('width', 'minwidth'):
                            widget.column(key, **{option: round(int(widget.column(key, option)) * factor / old_factor)})
                for child in widget.winfo_children():
                    columns(child)
            columns(root)
            scale_widgets(root)
            apply_style_scale(root)
        root.after(500, update)
    root.after(500, update)

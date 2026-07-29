"""Windows-only native dark title-bar support for PyQt5 top-level windows."""
from __future__ import annotations

import ctypes
import sys


def applydarktitlebar(widget) -> None:
    """Ask Windows 10/11 DWM to render this Qt top-level window's title bar dark.

    This is intentionally a no-op on non-Windows systems and on older Windows builds
    that do not support the immersive dark-mode DWM attribute.
    """
    if sys.platform != "win32":
        return
    try:
        hwnd = int(widget.winId())
        value = ctypes.c_int(1)
        dwmapi = ctypes.windll.dwmapi
        # Windows 10 builds differ: 20 is current; 19 is a compatibility fallback.
        for attribute in (20, 19):
            result = dwmapi.DwmSetWindowAttribute(
                ctypes.c_void_p(hwnd),
                ctypes.c_uint(attribute),
                ctypes.byref(value),
                ctypes.sizeof(value),
            )
            if result == 0:
                return
    except Exception:
        # Appearance enhancement only: never prevent the application from opening.
        return

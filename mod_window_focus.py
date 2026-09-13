"""
Windows 視窗焦點輔助：在 a300 查字時，把焦點從 Excel 搶回啟動程式的終端機
（含 WezTerm、Windows Terminal、PowerShell 等）。
"""

from __future__ import annotations

import ctypes
import logging
import os
import sys
import time
from ctypes import wintypes

try:
    import win32api
    import win32con
    import win32gui
    import win32process

    HAS_WIN32 = True
except ImportError:
    HAS_WIN32 = False

CONSOLE_TITLE_KEYWORDS = (
    "wezterm",
    "windows terminal",
    "alacritty",
    "windowsterminal",
    "python",
    "powershell",
    "pwsh",
    "cmd",
    "terminal",
    "piau-im",
    "vscode",
    "cursor",
)

CONSOLE_PROCESS_NAMES = {
    "wezterm-gui.exe",
    "wezterm.exe",
    "windowsterminal.exe",
    "alacritty.exe",
    "powershell.exe",
    "pwsh.exe",
    "cmd.exe",
}

EXCEL_PROCESS_NAMES = {
    "excel.exe",
}


def _is_valid_window(hwnd) -> bool:
    return bool(HAS_WIN32 and hwnd and win32gui.IsWindow(hwnd))


def _process_image_name(pid: int) -> str:
    """取得行程映像檔名（小寫），例如 wezterm-gui.exe。"""
    if not pid:
        return ""
    PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
    kernel32 = ctypes.windll.kernel32
    handle = kernel32.OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, pid)
    if not handle:
        return ""
    try:
        buf = ctypes.create_unicode_buffer(32768)
        size = wintypes.DWORD(len(buf))
        if kernel32.QueryFullProcessImageNameW(handle, 0, buf, ctypes.byref(size)):
            return os.path.basename(buf.value).lower()
        return ""
    finally:
        kernel32.CloseHandle(handle)


def _window_process_name(hwnd) -> str:
    try:
        _tid, pid = win32process.GetWindowThreadProcessId(hwnd)
        return _process_image_name(pid)
    except Exception:
        return ""


def _is_excel_window(hwnd, excel_hwnd=None) -> bool:
    if not _is_valid_window(hwnd):
        return False
    if excel_hwnd and hwnd == excel_hwnd:
        return True
    return _window_process_name(hwnd) in EXCEL_PROCESS_NAMES


def _is_console_window(hwnd, excel_hwnd=None) -> bool:
    if not _is_valid_window(hwnd) or _is_excel_window(hwnd, excel_hwnd):
        return False
    title = (win32gui.GetWindowText(hwnd) or "").lower()
    if any(keyword in title for keyword in CONSOLE_TITLE_KEYWORDS):
        return True
    return _window_process_name(hwnd) in CONSOLE_PROCESS_NAMES


def _enum_console_windows(excel_hwnd=None) -> list:
    windows = []

    def enum_handler(hwnd, result_list):
        if win32gui.IsWindowVisible(hwnd) and _is_console_window(hwnd, excel_hwnd):
            # 略過沒有標題的輔助視窗
            if win32gui.GetWindowText(hwnd):
                result_list.append(hwnd)

    if HAS_WIN32:
        win32gui.EnumWindows(enum_handler, windows)
    return windows


def capture_console_hwnd(excel_hwnd=None):
    """
    找出啟動本程式的終端機視窗。

    優先使用目前前景視窗（啟動 a300 當下通常就是 WezTerm），
    再退而依行程名稱／視窗標題搜尋。
    """
    if not HAS_WIN32:
        return None

    try:
        foreground = win32gui.GetForegroundWindow()
        if _is_valid_window(foreground) and not _is_excel_window(foreground, excel_hwnd):
            # 啟動當下的前景視窗就是 Terminal；即使標題不含舊關鍵字也一律採用
            return foreground
    except Exception:
        pass

    windows = _enum_console_windows(excel_hwnd)
    return windows[0] if windows else None


def force_foreground_window(hwnd) -> bool:
    """
    強制將指定視窗設為前景。

    Excel COM（activate／select）常會把焦點搶走；Windows 又限制背景程式
    呼叫 SetForegroundWindow。這裡用 AttachThreadInput、TOPMOST 切換與
    SwitchToThisWindow，盡量把焦點搶回終端機。
    """
    if not _is_valid_window(hwnd):
        return False

    try:
        user32 = ctypes.windll.user32

        if win32gui.IsIconic(hwnd):
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            time.sleep(0.15)

        if win32gui.GetForegroundWindow() == hwnd:
            return True

        foreground_hwnd = win32gui.GetForegroundWindow()
        foreground_tid, _ = win32process.GetWindowThreadProcessId(foreground_hwnd)
        target_tid, _ = win32process.GetWindowThreadProcessId(hwnd)
        current_tid = win32api.GetCurrentThreadId()

        attached = []
        for src_tid in (foreground_tid, current_tid):
            if src_tid and src_tid != target_tid:
                try:
                    win32process.AttachThreadInput(src_tid, target_tid, True)
                    attached.append(src_tid)
                except Exception as e:
                    logging.debug(f"AttachThreadInput 失敗: {e}")

        try:
            flags = win32con.SWP_NOMOVE | win32con.SWP_NOSIZE | win32con.SWP_SHOWWINDOW
            win32gui.SetWindowPos(hwnd, win32con.HWND_TOPMOST, 0, 0, 0, 0, flags)
            win32gui.SetWindowPos(hwnd, win32con.HWND_NOTOPMOST, 0, 0, 0, 0, flags)
            win32gui.ShowWindow(hwnd, win32con.SW_SHOW)
            win32gui.BringWindowToTop(hwnd)
            try:
                user32.SwitchToThisWindow(hwnd, True)
            except Exception:
                pass
            try:
                win32gui.SetForegroundWindow(hwnd)
            except Exception:
                pass
            try:
                win32gui.SetActiveWindow(hwnd)
            except Exception:
                pass

            if win32gui.GetForegroundWindow() != hwnd:
                # 最後手段：最小化再還原，通常能搶回前景
                win32gui.ShowWindow(hwnd, win32con.SW_MINIMIZE)
                time.sleep(0.05)
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                try:
                    win32gui.SetForegroundWindow(hwnd)
                except Exception:
                    pass
        finally:
            for src_tid in attached:
                try:
                    win32process.AttachThreadInput(src_tid, target_tid, False)
                except Exception as e:
                    logging.debug(f"DetachThreadInput 失敗: {e}")

        time.sleep(0.05)
        return win32gui.GetForegroundWindow() == hwnd
    except Exception as e:
        logging.debug(f"強制切換前景視窗失敗：{e}")
        return False


def activate_console_window(console_hwnd=None, excel_hwnd=None, quiet: bool = False) -> bool:
    """將終端機視窗切到前景，供使用者輸入讀音選項。"""
    if not HAS_WIN32:
        if not quiet:
            print("提示：無法自動切換到終端機視窗（需要 pywin32 套件）")
        return False

    hwnd = console_hwnd if _is_valid_window(console_hwnd) else None
    if not hwnd:
        hwnd = capture_console_hwnd(excel_hwnd=excel_hwnd)

    if not hwnd:
        print("提示：無法找到終端機視窗，請用滑鼠點一下 WezTerm／Terminal")
        return False

    ok = force_foreground_window(hwnd)
    if ok:
        if not quiet:
            print("✓ 已切換到終端機視窗")
            print("視窗焦點已正確設置")
    else:
        print("⚠️  無法自動把焦點切回終端機，請用滑鼠點一下 WezTerm／Terminal")

    try:
        sys.stdout.flush()
    except Exception:
        pass
    return ok

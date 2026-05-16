"""Windows-specific helpers for launching apps, URLs, clipboard and UI automation."""

from __future__ import annotations

import ctypes
import ctypes.wintypes
import os
import shutil
import subprocess
import time
import webbrowser
from pathlib import Path


def open_url(url: str) -> None:
    """Open a URL or Windows URI scheme with the default registered handler."""
    if not url:
        return
    try:
        os.startfile(url)  # type: ignore[attr-defined]
    except Exception:
        webbrowser.open(url)


def copy_to_clipboard(text: str) -> None:
    try:
        import tkinter as tk

        root = tk.Tk()
        root.withdraw()
        root.clipboard_clear()
        root.clipboard_append(text)
        root.update()
        root.destroy()
    except Exception:
        subprocess.run(
            ["powershell", "-NoProfile", "-Command", "Set-Clipboard -Value $input"],
            input=text,
            text=True,
            check=True,
            timeout=5,
        )


def press_key(key: str, delay: float = 0.0) -> tuple[bool, str]:
    try:
        import pyautogui

        if delay:
            time.sleep(delay)
        pyautogui.press(key)
        return True, "ok"
    except Exception as exc:
        return False, f"PyAutoGUI tus basimi yapamadi: {exc}"


def hotkey(*keys: str, delay: float = 0.0) -> tuple[bool, str]:
    try:
        import pyautogui

        if delay:
            time.sleep(delay)
        pyautogui.hotkey(*keys)
        return True, "ok"
    except Exception as exc:
        return False, f"PyAutoGUI kisayolu calistiramadi: {exc}"


def paste_text(text: str, delay: float = 0.0) -> tuple[bool, str]:
    try:
        copy_to_clipboard(text)
        return hotkey("ctrl", "v", delay=delay)
    except Exception as exc:
        return False, f"Pano veya yapistirma islemi basarisiz: {exc}"


def find_executable(names: list[str]) -> str | None:
    for name in names:
        found = shutil.which(name)
        if found:
            return found
    return None


def start_process(target: str) -> bool:
    if not target:
        return False
    try:
        os.startfile(target)  # type: ignore[attr-defined]
        return True
    except Exception:
        pass
    try:
        subprocess.Popen([target], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        return True
    except Exception:
        return False


def search_start_menu(shortcut_name: str) -> Path | None:
    roots = [
        Path(os.environ.get("APPDATA", "")) / "Microsoft" / "Windows" / "Start Menu" / "Programs",
        Path(os.environ.get("PROGRAMDATA", "")) / "Microsoft" / "Windows" / "Start Menu" / "Programs",
    ]
    target = shortcut_name.casefold()
    for root in roots:
        if not root.exists():
            continue
        for item in root.rglob("*.lnk"):
            stem = item.stem.casefold()
            if target == stem or target in stem:
                return item
    return None


def get_active_window_rect() -> tuple[int, int, int, int] | None:
    try:
        user32 = ctypes.windll.user32
        hwnd = user32.GetForegroundWindow()
        if not hwnd:
            return None
        rect = ctypes.wintypes.RECT()
        if not user32.GetWindowRect(hwnd, ctypes.byref(rect)):
            return None
        if rect.right <= rect.left or rect.bottom <= rect.top:
            return None
        return rect.left, rect.top, rect.right, rect.bottom
    except Exception:
        return None


def get_active_window_title() -> str:
    try:
        user32 = ctypes.windll.user32
        hwnd = user32.GetForegroundWindow()
        length = user32.GetWindowTextLengthW(hwnd)
        buffer = ctypes.create_unicode_buffer(length + 1)
        user32.GetWindowTextW(hwnd, buffer, length + 1)
        return buffer.value.strip()
    except Exception:
        return ""

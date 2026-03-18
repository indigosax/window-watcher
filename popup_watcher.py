"""
popup_watcher.py
================
Monitors ALL new windows appearing on screen. When a new window is detected:
  - Captures a screenshot
  - Logs the window title, process name, PID, and timestamp
  - Saves screenshots to a timestamped folder

Requirements:
    pip install pywin32 pillow psutil

Usage:
    python popup_watcher.py

Output:
    - Screenshots saved to: C:\PopupWatcher\screenshots\
    - Log file saved to:    C:\PopupWatcher\popup_log.txt

Press Ctrl+C to stop.
"""

import time
import os
import ctypes
import ctypes.wintypes
import logging
from datetime import datetime

try:
    import win32gui
    import win32process
    import win32con
    import win32api
    import psutil
    from PIL import ImageGrab
except ImportError:
    print("ERROR: Missing dependencies. Please run:")
    print("    pip install pywin32 pillow psutil")
    input("Press Enter to exit...")
    exit(1)

# ── Configuration ─────────────────────────────────────────────────────────────
OUTPUT_DIR      = r"C:\PopupWatcher"
SCREENSHOT_DIR  = os.path.join(OUTPUT_DIR, "screenshots")
LOG_FILE        = os.path.join(OUTPUT_DIR, "popup_log.txt")
POLL_INTERVAL   = 0.25   # seconds between scans (lower = more sensitive, more CPU)
MIN_WINDOW_W    = 100    # ignore tiny windows smaller than this (pixels)
MIN_WINDOW_H    = 50

# Window titles/classes to ignore (noisy system windows)
IGNORE_TITLES = {
    "", "Program Manager", "Windows Input Experience",
    "Windows Shell Experience Host", "Search", "Task Switching",
    "Cortana", "Microsoft Text Input Application",
    "NVIDIA GeForce Overlay", "NVIDIA GeForce Overlay DT",
    "Default IME", "MSCTFIME UI",
}

IGNORE_CLASSES = {
    "Shell_TrayWnd", "DV2ControlHost", "MsgrIMEWindowClass",
    "SysShadow", "Button", "tooltips_class32",
}
# ─────────────────────────────────────────────────────────────────────────────

os.makedirs(SCREENSHOT_DIR, exist_ok=True)

logging.basicConfig(
    filename=LOG_FILE,
    level=logging.INFO,
    format="%(asctime)s | %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)

def get_all_windows():
    """Return a dict of {hwnd: title} for all visible top-level windows."""
    windows = {}

    def callback(hwnd, _):
        if not win32gui.IsWindowVisible(hwnd):
            return
        title = win32gui.GetWindowText(hwnd)
        cls   = win32gui.GetClassName(hwnd)
        if title in IGNORE_TITLES or cls in IGNORE_CLASSES:
            return
        try:
            rect = win32gui.GetWindowRect(hwnd)
            w = rect[2] - rect[0]
            h = rect[3] - rect[1]
            if w < MIN_WINDOW_W or h < MIN_WINDOW_H:
                return
        except Exception:
            return
        windows[hwnd] = title

    win32gui.EnumWindows(callback, None)
    return windows


def get_process_info(hwnd):
    """Get process name and PID for a window handle."""
    try:
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        proc = psutil.Process(pid)
        return proc.name(), pid, proc.exe()
    except Exception:
        return "Unknown", 0, "Unknown"


def capture_screenshot(hwnd, label):
    """Capture full screen screenshot and save it."""
    ts = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    safe_label = "".join(c if c.isalnum() or c in " _-" else "_" for c in label)[:50]
    filename = f"{ts}_{safe_label}.png"
    filepath = os.path.join(SCREENSHOT_DIR, filename)
    try:
        img = ImageGrab.grab()
        img.save(filepath)
        return filepath
    except Exception as e:
        return f"Screenshot failed: {e}"


def capture_window_screenshot(hwnd, label):
    """Try to capture just the window region."""
    try:
        rect = win32gui.GetWindowRect(hwnd)
        ts = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
        safe_label = "".join(c if c.isalnum() or c in " _-" else "_" for c in label)[:50]
        filename = f"{ts}_{safe_label}.png"
        filepath = os.path.join(SCREENSHOT_DIR, filename)
        img = ImageGrab.grab(bbox=rect)
        img.save(filepath)
        return filepath
    except Exception:
        # Fall back to full screen
        return capture_screenshot(hwnd, label)


def log_and_print(msg):
    print(msg)
    logging.info(msg)


def main():
    print("=" * 60)
    print("  Popup Watcher — Running")
    print(f"  Screenshots → {SCREENSHOT_DIR}")
    print(f"  Log file    → {LOG_FILE}")
    print("  Press Ctrl+C to stop.")
    print("=" * 60)

    known_windows = get_all_windows()
    log_and_print(f"Started. Tracking {len(known_windows)} existing windows.")

    try:
        while True:
            time.sleep(POLL_INTERVAL)
            current_windows = get_all_windows()

            # Detect NEW windows
            new_hwnds = set(current_windows.keys()) - set(known_windows.keys())

            for hwnd in new_hwnds:
                title = current_windows.get(hwnd, "<no title>")
                proc_name, pid, exe_path = get_process_info(hwnd)

                # Capture screenshot immediately
                screenshot_path = capture_window_screenshot(hwnd, title or proc_name)

                msg = (
                    f"NEW WINDOW DETECTED\n"
                    f"  Title   : {title!r}\n"
                    f"  Process : {proc_name}  (PID {pid})\n"
                    f"  EXE     : {exe_path}\n"
                    f"  Screenshot: {screenshot_path}"
                )
                log_and_print(msg)
                print("-" * 60)

            # Detect CLOSED windows (optional info)
            closed_hwnds = set(known_windows.keys()) - set(current_windows.keys())
            for hwnd in closed_hwnds:
                title = known_windows.get(hwnd, "<no title>")
                if title:
                    log_and_print(f"CLOSED: {title!r}")

            known_windows = current_windows

    except KeyboardInterrupt:
        print("\nStopped. Check your log and screenshots at:")
        print(f"  {OUTPUT_DIR}")


if __name__ == "__main__":
    main()

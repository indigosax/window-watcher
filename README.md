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

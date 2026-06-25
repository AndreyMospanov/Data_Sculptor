from __future__ import annotations

import os
import sys
import traceback
from pathlib import Path


def show_startup_error(message: str) -> None:
    try:
        import tkinter as tk
        from tkinter import messagebox

        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("Data Sculptor startup error", message)
        root.destroy()
    except Exception:
        # If tkinter itself is missing or broken, keep a text log next to the launcher.
        log_path = Path(__file__).with_name("data_sculptor_gui_error.log")
        log_path.write_text(message, encoding="utf-8")


def main() -> int:
    launcher_dir = Path(__file__).resolve().parent
    os.chdir(launcher_dir)
    sys.path.insert(0, str(launcher_dir))

    try:
        from core.gui import main as gui_main
    except Exception:
        show_startup_error(
            "Could not import Data Sculptor GUI.\n\n"
            "Try running this command from the application folder:\n"
            "python -m core.gui\n\n"
            "Details:\n"
            f"{traceback.format_exc()}"
        )
        return 1

    try:
        return gui_main()
    except Exception:
        show_startup_error(
            "Data Sculptor crashed during startup.\n\n"
            "Details:\n"
            f"{traceback.format_exc()}"
        )
        return 1


raise SystemExit(main())

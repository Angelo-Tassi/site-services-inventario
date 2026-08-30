#!/usr/bin/env python3
"""Avvio dell'applicazione Inventario dispositivi."""

import sys

if __name__ == "__main__":
    try:
        import openpyxl  # noqa: F401
    except ImportError:
        sys.stderr.write(
            "Manca la libreria openpyxl.\n"
            "Installala con:  pip install openpyxl\n")
        try:
            import tkinter.messagebox as mb, tkinter as tk
            root = tk.Tk(); root.withdraw()
            mb.showerror("Componente mancante",
                         "Manca la libreria openpyxl.\n\n"
                         "Apri il Prompt dei comandi ed esegui:\n"
                         "pip install openpyxl")
        except Exception:
            pass
        sys.exit(2)

    from inventario.ui import main
    sys.exit(main())

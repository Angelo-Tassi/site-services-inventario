#!/bin/bash
# Avvio su macOS / Linux: doppio clic su questo file, oppure ./Avvia\ Inventario.command
cd "$(dirname "$0")" || exit 1
if [ -x ".venv/bin/python" ]; then
    exec .venv/bin/python Inventario.py
fi
exec python3 Inventario.py

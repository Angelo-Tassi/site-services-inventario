#!/usr/bin/env python3
"""Rigenera i file di prova in Collaudo/.

Uso:  .venv/bin/python tests/genera_file_di_prova.py
"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

RADICE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CARTELLA = os.path.join(RADICE, "Collaudo")

INTESTAZIONI = ["Asset Tag", "Tipo", "Modello", "Numero di serie", "Stato", "Note"]

LAPTOP = ["Lenovo ThinkPad T14 Gen 4", "Lenovo ThinkPad T14 Gen 5"]
TABLET = ["Dell Latitude 7320 Detachable", "Dell Latitude 7230 Rugged Extreme"]

STANZE = [
    ("SITE SERVICES BAU", "BAU", 100),
    ("DIGITAL KIOSK", "KSK", 200),
    ("MAGAZZINO DISASTER RECOVERY", "DRC", 300),
]

STATI = ["Disponibile", "Disponibile", "Disponibile", "In attesa ritiro",
         "Guasto in attesa tecnico", "Da rebuildare", "Controllare",
         "Disponibile", "Disponibile", "Disponibile"]

NOTE = ["", "Postazione 1", "", "Batteria da sostituire", "", "Con custodia",
        "", "In verifica", "", "Scorta"]


def _dispositivo(prefisso, base, indice):
    """Sette laptop e tre tablet per stanza, con seriali plausibili."""
    numero = base + indice
    if indice <= 7:
        modello = LAPTOP[indice % len(LAPTOP)]
        seriale = "PF%d%s%02d" % (4 + indice % 2, prefisso[:2], numero % 100)
        tipo = "Laptop"
    else:
        modello = TABLET[indice % len(TABLET)]
        seriale = "%dH%s%03d" % (4 + indice % 5, prefisso[:1], numero % 1000)
        tipo = "Tablet"
    return ["IT-%s-%03d" % (prefisso, numero), tipo, modello, seriale,
            STATI[indice - 1], NOTE[indice - 1]]


def _intesta(ws):
    ws.append(INTESTAZIONI)
    for cella in ws[1]:
        cella.font = Font(bold=True, color="FFFFFF")
        cella.fill = PatternFill("solid", fgColor="1F4E79")
        cella.alignment = Alignment(vertical="center")
    ws.freeze_panes = "A2"
    for lettera, larghezza in zip("ABCDEF", (16, 12, 36, 20, 24, 30)):
        ws.column_dimensions[lettera].width = larghezza


def _separatore(ws, testo):
    riga = ws.max_row + 1
    cella = ws.cell(row=riga, column=1, value=testo)
    cella.font = Font(bold=True, color="1F4E79")
    cella.fill = PatternFill("solid", fgColor="DCE6F1")
    ws.merge_cells(start_row=riga, start_column=1, end_row=riga,
                   end_column=len(INTESTAZIONI))


def inventario_completo(percorso):
    """Trenta dispositivi, dieci per stanza, divisi dai separatori."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Inventario"
    _intesta(ws)
    for tag, prefisso, base in STANZE:
        _separatore(ws, tag)
        for indice in range(1, 11):
            ws.append(_dispositivo(prefisso, base, indice))
    wb.save(percorso)
    wb.close()
    return percorso


def inventario_con_difetti(percorso):
    """Lo stesso file, con dentro i casi che il programma deve segnalare."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Inventario"
    ws.append(INTESTAZIONI + ["IMEI", "Costo", "Fornitore", "Centro di costo"])
    for cella in ws[1]:
        cella.font = Font(bold=True, color="FFFFFF")
        cella.fill = PatternFill("solid", fgColor="1F4E79")
    _separatore(ws, "BAU")
    ws.append(["IT-BAU-901", "Laptop", "Lenovo ThinkPad T14 Gen 5", "PF4BAU01",
               "Disponibile", "riga regolare", "", 1200, "Dell Italia", "CC-01"])
    ws.append(["IT-BAU-902", "Laptop", "", "PF4BAU02", "Disponibile",
               "senza modello", "", 1150, "Dell Italia", "CC-01"])
    ws.append([None] * 10)
    _separatore(ws, "KIOSK")
    ws.append(["IT-KSK-903", "Tablet", "Dell Latitude 7320 Detachable", "8HK903",
               "Controllare", "", "", 900, "Dell Italia", "CC-02"])
    ws.append(["", "Laptop", "senza identificativo", "", "", "verra' scartata",
               "", 800, "Dell Italia", "CC-02"])
    ws.append(["", "Iphone", "Apple iPhone 14", "", "", "verra' ignorato",
               "356938035643809", 1000, "Apple", "CC-03"])
    wb.save(percorso)
    wb.close()
    return percorso


def main():
    if not os.path.isdir(CARTELLA):
        os.makedirs(CARTELLA)
    uno = inventario_completo(os.path.join(CARTELLA, "Inventario_di_prova.xlsx"))
    due = inventario_con_difetti(os.path.join(CARTELLA, "Inventario_di_prova_con_difetti.xlsx"))
    for percorso in (uno, due):
        print("scritto:", os.path.relpath(percorso, RADICE))


if __name__ == "__main__":
    main()

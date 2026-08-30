#!/usr/bin/env python3
"""Rigenera i dati dimostrativi in Esempio/.

Sono dati finti, pensati per far vedere il programma a chi lo apre la prima
volta: tredici dispositivi nelle tre stanze, con due prestiti in corso nel
Digital Kiosk. Non e' l'inventario vero, che vive fuori dal repository.

Uso:  .venv/bin/python tests/genera_esempio.py
"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from inventario import config
from inventario.store import InventoryStore, new_item

RADICE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CARTELLA = os.path.join(RADICE, "Esempio")

BAU = "Site Services BAU"
KIOSK = "Digital Kiosk"
DR = "Magazzino Disaster Recovery"

DEMO = [
    ("IT-0101", "Laptop", "Lenovo ThinkPad T14 Gen 4", "PF4A1B2C", BAU, "Postazione reception"),
    ("IT-0102", "Laptop", "Lenovo ThinkPad T14 Gen 4", "PF4A1B7D", BAU, ""),
    ("IT-0103", "Laptop", "Lenovo ThinkPad T14 Gen 5", "PF5K9M3E", BAU, "Batteria da sostituire"),
    ("IT-0104", "Tablet", "Dell Latitude 7320 Detachable", "8H2KLM3", BAU, "Con tastiera e pennino"),
    ("IT-0105", "Tablet", "Dell Latitude 7230 Rugged Extreme", "4RT9WQ2", BAU, ""),
    ("IT-0106", "Laptop", "Lenovo ThinkPad T14 Gen 5", "PF5K9M8F", KIOSK, "Postazione kiosk 1"),
    ("IT-0107", "Laptop", "Lenovo ThinkPad T14 Gen 4", "PF4A2C1G", KIOSK, ""),
    ("IT-0108", "Laptop", "Lenovo ThinkPad T14 Gen 5", "PF5L4N2H", KIOSK, "Postazione kiosk 2"),
    ("IT-0109", "Tablet", "Dell Latitude 7320 Detachable", "8H2KLP9", KIOSK, ""),
    ("IT-0110", "Tablet", "Dell Latitude 7230 Rugged Extreme", "4RT9WX7", KIOSK, "Con custodia"),
    ("DR-0201", "Laptop", "Lenovo ThinkPad T14 Gen 4", "PF4B7T1J", DR, "Scorta sigillata"),
    ("DR-0202", "Laptop", "Lenovo ThinkPad T14 Gen 5", "PF5M2P4K", DR, "Scorta sigillata"),
    ("DR-0203", "Tablet", "Dell Latitude 7320 Detachable", "8H3NQR5", DR, "Kit continuita' operativa"),
]


def main():
    if not os.path.isdir(CARTELLA):
        os.makedirs(CARTELLA)
    percorso = os.path.join(CARTELLA, "Inventario.xlsx")
    if os.path.exists(percorso):
        os.remove(percorso)
    archivio = InventoryStore(percorso, iphone_room=BAU)
    archivio.create_if_missing()
    for riga in DEMO:
        archivio.add(new_item(*riga))
    archivio.lend("IT-0107", "Marco Bianchi")
    archivio.lend("IT-0110", "Elena Rossi")
    config.save_shared_config(percorso, {
        "rooms": [BAU, KIOSK, DR],
        "types": ["Laptop", "Tablet", "Iphone"],
        "loan_rooms": [KIOSK],
        "iphone_room": BAU,
    })
    archivio.load()
    print("scritti %d dispositivi in %s"
          % (len(archivio.items), os.path.relpath(percorso, RADICE)))
    return 0


if __name__ == "__main__":
    sys.exit(main())

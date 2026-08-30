"""Gli iPhone non spariscono mai, per nessuna strada distruttiva.

Le vie che cancellano dati sono tre: il reset, l'importazione in sostituzione su
tutto l'inventario e quella su una singola stanza. Nessuna deve toccare i
telefoni, che non si possono reimportare da un file.
"""
import os, sys, tempfile
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import fixture
from datetime import datetime, timedelta
from inventario.store import (BloccoConservazione, BloccoIphoneNonSpedito,
                              InventoryStore, new_item, puo_essere_eliminato)

BAU, KIOSK, DR = fixture.BAU, fixture.KIOSK, fixture.DR
d = tempfile.mkdtemp()

def inventario():
    """Tre telefoni in stati diversi, piu' laptop e tablet."""
    s = InventoryStore(os.path.join(tempfile.mkdtemp(), "Inventario.xlsx"), iphone_room=BAU)
    s.create_if_missing()
    for tag, stanza in (("IT-0101", BAU), ("IT-0106", KIOSK), ("DR-0201", DR)):
        s.add(new_item(tag, "Laptop", "T14", "SN" + tag, stanza))
    s.add(new_item(tipo="Iphone", modello="iPhone 14", imei="351000000000001",
                   restituito_da="Mai spedito"))
    s.add(new_item(tipo="Iphone", modello="iPhone 13", imei="351000000000002",
                   restituito_da="Spedito ieri"))
    s.ship("351000000000002")
    vecchio = new_item(tipo="Iphone", modello="iPhone 12", imei="351000000000003",
                       restituito_da="Spedito da un anno")
    vecchio["spedito_il"] = (datetime.now() - timedelta(days=400)).strftime("%d/%m/%Y %H:%M")
    s.add(vecchio)
    s.load()
    return s

TELEFONI = {"351000000000001", "351000000000002", "351000000000003"}

def telefoni(s):
    return {i["asset_tag"] for i in s.items if i["tipo"].lower() == "iphone"}

# ---- di partenza: sei dispositivi, tre telefoni
s = inventario()
assert len(s.items) == 6 and telefoni(s) == TELEFONI
# il terzo e' spedito da oltre tre mesi: sarebbe eliminabile a mano
assert puo_essere_eliminato(s._read()[0] if False else
                            [i for i in s.items if i["asset_tag"] == "351000000000003"][0])[0]

# ---- 1. reset completo
eliminati, tenuti, copia = s.reset()
s.load()
assert telefoni(s) == TELEFONI, "il reset non tocca nessun iPhone"
assert eliminati == 3 and tenuti == 3, (eliminati, tenuti)
assert os.path.exists(copia)

# ---- 2. importazione in sostituzione su tutto l'inventario
s = inventario()
nuovi = [new_item("IT-0900", "Laptop", "T14 Gen 5", "PF900", KIOSK)]
esito = s.import_items(nuovi, "replace", None)
s.load()
assert telefoni(s) == TELEFONI, "la sostituzione totale non tocca gli iPhone"
assert esito["eliminati"] == 3
assert len(s.items) == 4

# ---- 3. importazione in sostituzione sulla stanza degli iPhone
s = inventario()
esito = s.import_items([new_item("IT-0901", "Laptop", "T14", "PF901", BAU)], "replace", BAU)
s.load()
assert telefoni(s) == TELEFONI, "nemmeno sostituendo proprio la loro stanza"
assert esito["eliminati"] == 1, esito     # solo il laptop che stava in BAU

# ---- 4. due sostituzioni di fila
s.import_items([new_item("IT-0902", "Laptop", "T14", "PF902", BAU)], "replace", BAU)
s.import_items([], "replace", None)
s.load()
assert telefoni(s) == TELEFONI, "nemmeno svuotando tutto con un file vuoto"
assert {i["asset_tag"] for i in s.items} == TELEFONI

# ---- l'eliminazione a mano resta regolata dalle sue condizioni
s = inventario()
try:
    s.delete(["351000000000001"]); raise SystemExit("iPhone mai spedito eliminato")
except BloccoIphoneNonSpedito:
    pass
try:
    s.delete(["351000000000002"]); raise SystemExit("iPhone spedito ieri eliminato")
except BloccoConservazione:
    pass
assert s.delete(["351000000000003"]) == 1, "dopo tre mesi si elimina a mano"
s.load()
assert telefoni(s) == {"351000000000001", "351000000000002"}
print("IPHONE PROTETTI OK")

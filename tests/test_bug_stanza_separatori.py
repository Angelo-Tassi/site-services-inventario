"""Il caso segnalato: importando 'Una sola stanza' dalla pagina principale, i
separatori venivano ignorati e finiva tutto nella stanza scelta.

Le due strade - la finestra generale e il pulsante dentro la stanza - devono
comportarsi in modo identico.
"""
import os, sys, tempfile
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import fixture
from openpyxl import Workbook
from inventario.store import InventoryStore, rows_from_workbook
from inventario.ui import seleziona_per_stanza

RADICE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
COMPLETO = os.path.join(RADICE, "Collaudo", "Inventario_di_prova.xlsx")
BAU, KIOSK, DR = fixture.BAU, fixture.KIOSK, fixture.DR
STANZE = [BAU, KIOSK, DR]
d = tempfile.mkdtemp()

# ---- il caso del bug: file con i separatori, importazione in una sola stanza
items, esito = rows_from_workbook(COMPLETO, STANZE)
assert len(items) == 30 and esito["stanze_trovate"] == STANZE

miei, scartati, regola = seleziona_per_stanza(items, esito, KIOSK)
assert regola == "separatori", regola
assert len(miei) == 10, len(miei)
assert scartati == 20, scartati
assert all(i["stanza"] == KIOSK for i in miei)
assert not any(i["asset_tag"].startswith("IT-BAU") for i in miei)
assert not any(i["asset_tag"].startswith("IT-DRC") for i in miei)

# lo stesso vale per ogni stanza
for stanza, prefisso in ((BAU, "IT-BAU"), (DR, "IT-DRC")):
    righe, fuori, come = seleziona_per_stanza(items, esito, stanza)
    assert come == "separatori" and len(righe) == 10 and fuori == 20
    assert all(i["asset_tag"].startswith(prefisso) for i in righe), stanza

# e l'inventario che ne risulta contiene solo quella stanza
s = InventoryStore(os.path.join(d, "Inventario.xlsx"), iphone_room=BAU)
s.create_if_missing()
s.import_items(miei, "merge", KIOSK)
s.load()
assert len(s.items) == 10 and all(i["stanza"] == KIOSK for i in s.items)

# ---- il foglio senza separatori: la scelta dell'utente vale per tutto
def foglio(nome, righe, con_separatori=False):
    p = os.path.join(d, nome + ".xlsx")
    wb = Workbook(); ws = wb.active
    ws.append(["Asset Tag", "Tipo", "Modello", "Numero di serie"])
    for r in righe:
        ws.append(r)
    wb.save(p); wb.close()
    return p

piatto = foglio("piatto", [["IT-0901", "Laptop", "T14 Gen 5", "PF901"],
                           ["IT-0902", "Tablet", "Dell Latitude", "8H902"]])
items2, esito2 = rows_from_workbook(piatto, STANZE)
assert esito2["stanze_trovate"] == []
righe, fuori, come = seleziona_per_stanza(items2, esito2, DR)
assert come == "tutte" and len(righe) == 2 and fuori == 0
s.import_items(righe, "merge", DR)
s.load()
assert sum(1 for i in s.items if i["stanza"] == DR) == 2

# ---- il foglio dichiara stanze, ma non quella richiesta: non si importa niente
altra = foglio("altra", [["BAU"], ["IT-0903", "Laptop", "T14", "PF903"]])
items3, esito3 = rows_from_workbook(altra, STANZE)
assert esito3["stanze_trovate"] == [BAU]
righe, fuori, come = seleziona_per_stanza(items3, esito3, KIOSK)
assert come == "mancante" and righe is None, (come, righe)

prima = len(s.items)
s.load()
assert len(s.items) == prima, "niente deve essere cambiato"

# ---- separatore presente ma sezione vuota
vuota = foglio("vuota", [["KIOSK"], ["BAU"], ["IT-0904", "Laptop", "T14", "PF904"]])
items4, esito4 = rows_from_workbook(vuota, STANZE)
righe, fuori, come = seleziona_per_stanza(items4, esito4, KIOSK)
assert come == "separatori" and righe == [], (come, righe)
print("BUG STANZA SEPARATORI OK")

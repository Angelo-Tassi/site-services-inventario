"""Il pulsante 'Importa i dati di questa stanza', provato sui file di Collaudo/."""
import os, sys, tempfile
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import fixture
from openpyxl import Workbook, load_workbook
from inventario.store import InventoryStore, new_item, rows_from_workbook
from inventario.ui import App, ImportOptionsDialog

RADICE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
COMPLETO = os.path.join(RADICE, "Collaudo", "Inventario_di_prova.xlsx")
BAU, KIOSK, DR = fixture.BAU, fixture.KIOSK, fixture.DR
STANZE = [BAU, KIOSK, DR]
d = tempfile.mkdtemp()

app = App(fixture.build()); app._initial_load()

# ---- la finestra a stanza fissa non chiede l'ambito, solo la modalita'
dlg = ImportOptionsDialog(app, STANZE, stanza_fissa=KIOSK)
assert dlg.combo is None, "niente tendina: la stanza e' gia' decisa"
assert dlg.var_stanza.get() == KIOSK and dlg.var_ambito.get() == "stanza"
assert dlg.title() == "Importa in %s" % KIOSK
dlg.var_mode.set("replace"); dlg._ok()
assert dlg.result == {"stanza": KIOSK, "mode": "replace"}, dlg.result

# ---- il file di prova dichiara tutte e tre le stanze
items, esito = rows_from_workbook(COMPLETO, STANZE)
assert esito["stanze_trovate"] == [BAU, KIOSK, DR], esito["stanze_trovate"]
assert len(items) == 30

# ---- si prende solo la sezione della stanza, il resto si scarta
for stanza in STANZE:
    miei = [i for i in items if i["stanza"] == stanza]
    assert len(miei) == 10, (stanza, len(miei))
    assert len(items) - len(miei) == 20

s = InventoryStore(os.path.join(d, "Inventario.xlsx"), iphone_room=BAU)
s.create_if_missing()
miei = [i for i in items if i["stanza"] == KIOSK]
r = s.import_items(miei, "merge", KIOSK)
assert (r["aggiunti"], r["gia_presenti"]) == (10, []), r
s.load()
assert len(s.items) == 10 and all(i["stanza"] == KIOSK for i in s.items)
assert not any(i["asset_tag"].startswith("IT-BAU") for i in s.items), \
    "le righe delle altre stanze non devono entrare"

# ---- rifare la stessa importazione non duplica: sono tutti gia' in inventario,
# quindi non entra niente e si dice dove stanno quelli che ci sono gia'
r = s.import_items(miei, "merge", KIOSK)
assert r["aggiunti"] == 0, r
assert len(r["gia_presenti"]) == 10, r["gia_presenti"]
assert all(v["stanza"] == KIOSK for v in r["gia_presenti"]), r["gia_presenti"]
s.load(); assert len(s.items) == 10

# ---- in sostituzione tocca solo quella stanza
s.import_items([i for i in items if i["stanza"] == BAU], "merge", BAU)
s.load(); assert len(s.items) == 20
r = s.import_items(miei, "replace", KIOSK)
assert r["eliminati"] == 10, r
s.load()
conteggi = {}
for i in s.items:
    conteggi[i["stanza"]] = conteggi.get(i["stanza"], 0) + 1
assert conteggi == {BAU: 10, KIOSK: 10}, conteggi

# ---- senza la riga della stanza non si importa niente
def foglio(nome, righe):
    p = os.path.join(d, nome + ".xlsx")
    wb = Workbook(); ws = wb.active
    ws.append(["Asset Tag", "Tipo", "Modello", "Numero di serie"])
    for r in righe:
        ws.append(r)
    wb.save(p); wb.close()
    return p

senza = foglio("senza_separatore",
               [["IT-0901", "Laptop", "T14 Gen 5", "PF901"],
                ["IT-0902", "Laptop", "T14 Gen 4", "PF902"]])
items2, esito2 = rows_from_workbook(senza, STANZE)
assert esito2["stanze_trovate"] == [], esito2
assert KIOSK not in esito2["stanze_trovate"], "il programma deve rifiutare"
assert len(items2) == 2, "le righe ci sono, ma non si sa di chi siano"

# ---- con il separatore di un'altra stanza soltanto
altra = foglio("altra_stanza", [["BAU"], ["IT-0903", "Laptop", "T14", "PF903"]])
items3, esito3 = rows_from_workbook(altra, STANZE)
assert esito3["stanze_trovate"] == [BAU]
assert KIOSK not in esito3["stanze_trovate"]
assert [i for i in items3 if i["stanza"] == KIOSK] == []

# ---- separatore presente ma sezione vuota
vuota = foglio("sezione_vuota",
               [["DIGITAL KIOSK"], ["BAU"], ["IT-0904", "Laptop", "T14", "PF904"]])
items4, esito4 = rows_from_workbook(vuota, STANZE)
assert KIOSK in esito4["stanze_trovate"], esito4
assert [i for i in items4 if i["stanza"] == KIOSK] == [], "niente da caricare"

# ---- il separatore si scrive anche in forma breve
breve = foglio("forma_breve", [["kiosk"], ["IT-0905", "Laptop", "T14", "PF905"]])
items5, esito5 = rows_from_workbook(breve, STANZE)
assert esito5["stanze_trovate"] == [KIOSK], esito5
assert len([i for i in items5 if i["stanza"] == KIOSK]) == 1

# ---- e un file esportato da una stanza si reimporta in quella stanza
from inventario import excel_io
uscita = os.path.join(d, "export_kiosk.xlsx")
excel_io.export([i for i in s.items if i["stanza"] == KIOSK], uscita,
                rooms=[KIOSK], titolo=KIOSK)
items6, esito6 = rows_from_workbook(uscita, STANZE)
assert len([i for i in items6 if i["stanza"] == KIOSK]) == 10, esito6
app.destroy()
print("IMPORT STANZA OK")

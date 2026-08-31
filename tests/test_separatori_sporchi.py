"""Separatori di stanza riconosciuti anche in fogli poco ordinati.

Il caso vero: dopo un'importazione le stanze restano vuote e i nomi delle
stanze compaiono in elenco come se fossero dispositivi. Succede quando la riga
separatore non viene riconosciuta, e da li' in poi nessuna riga prende la
stanza.
"""
import os, sys, tempfile
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import fixture
from openpyxl import Workbook
from inventario.store import rows_from_workbook, separatore_con_avanzi, tag_stanze
from inventario.ui import ImportDialog

BAU, KIOSK = fixture.BAU, fixture.KIOSK
STANZE = [BAU, KIOSK, fixture.DR]
TAGS = tag_stanze(STANZE)

def scrivi(righe):
    wb = Workbook(); ws = wb.active; ws.title = "Inventario"
    for r in righe:
        ws.append(r)
    p = os.path.join(tempfile.mkdtemp(), "prova.xlsx")
    wb.save(p); wb.close()
    return p

INTESTAZIONI = ["Asset Tag", "Tipo", "Modello", "Numero di serie", "Stato", "Note"]

# ---- la forma pulita continua a funzionare
pulito = scrivi([INTESTAZIONI,
                 ["SITE SERVICES BAU", None, None, None, None, None],
                 ["IT-1", "Laptop", "T14 Gen 5", "PF1", "Disponibile", None],
                 ["KIOSK", None, None, None, None, None],
                 ["IT-2", "Tablet", "Dell 7320", "8H2", "Disponibile", None]])
items, esito = rows_from_workbook(pulito, STANZE)
assert [i["stanza"] for i in items] == [BAU, KIOSK], items
assert esito["stanze_trovate"] == [BAU, KIOSK]

# ---- separatore con celle di troppo: prima diventava un dispositivo
sporco = scrivi([INTESTAZIONI,
                 ["SITE SERVICES BAU", None, None, None, None, "sezione uffici"],
                 ["IT-1", "Laptop", "T14 Gen 5", "PF1", "Disponibile", None],
                 ["KIOSK", None, None, None, None, "aggiornato a marzo"],
                 ["IT-2", "Tablet", "Dell 7320", "8H2", "Disponibile", None]])
items, esito = rows_from_workbook(sporco, STANZE)
assert len(items) == 2, [i["asset_tag"] for i in items]
assert [i["asset_tag"] for i in items] == ["IT-1", "IT-2"]
assert [i["stanza"] for i in items] == [BAU, KIOSK], items
assert esito["stanze_trovate"] == [BAU, KIOSK]

# ---- un dispositivo vero non viene mai scambiato per un separatore
mapping = {0: "asset_tag", 1: "tipo", 2: "modello", 3: "seriale", 4: "stato"}
assert separatore_con_avanzi(("KIOSK", "Laptop", "T14", "PF9", "Disponibile"),
                             mapping, TAGS) is None
assert separatore_con_avanzi(("KIOSK", None, None, None, None), mapping, TAGS) == KIOSK
assert separatore_con_avanzi(("IT-9", None, None, None, None), mapping, TAGS) is None

# ---- senza nessun separatore l'utente viene avvisato prima di importare
senza = scrivi([INTESTAZIONI,
                ["IT-1", "Laptop", "T14 Gen 5", "PF1", "Disponibile", None],
                ["IT-2", "Tablet", "Dell 7320", "8H2", "Disponibile", None]])
items, esito = rows_from_workbook(senza, STANZE)
assert esito["stanze_trovate"] == []
avvisi = ImportDialog._avvertenze(esito, {"stanza": None, "mode": "merge"}, len(items))
assert any("SENZA STANZA" in a for a in avvisi), avvisi
assert any("2 dispositivi" in a for a in avvisi), avvisi

# ---- ma non si avvisa quando le stanze ci sono, o quando si importa in una sola
avvisi = ImportDialog._avvertenze({"stanze_trovate": [BAU]},
                                  {"stanza": None, "mode": "merge"}, 5)
assert not [a for a in avvisi if "SENZA STANZA" in a]
avvisi = ImportDialog._avvertenze({"stanze_trovate": []},
                                  {"stanza": KIOSK, "mode": "merge"}, 5)
assert not [a for a in avvisi if "SENZA STANZA" in a]

print("SEPARATORI OK")

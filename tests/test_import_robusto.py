"""Importazione di file con colonne in piu', mancanti o non riconosciute."""
import os, sys, tempfile
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import fixture
from openpyxl import Workbook
from inventario.store import InventoryError, InventoryStore, map_headers, rows_from_workbook
from inventario.ui import ImportDialog

STANZE = [fixture.BAU, fixture.KIOSK, fixture.DR]
d = tempfile.mkdtemp()

def foglio(nome, intestazioni, righe):
    p = os.path.join(d, nome + ".xlsx")
    wb = Workbook(); ws = wb.active
    ws.append(intestazioni)
    for r in righe:
        ws.append(r)
    wb.save(p); wb.close()
    return p

# ---- colonne in piu': ignorate, ma segnalate
p = foglio("extra", ["Asset Tag", "Tipo", "Modello", "Numero di serie",
                     "Costo", "Fornitore", "Centro di costo"],
           [["IT-0101", "Laptop", "T14 Gen 4", "PF4A1B2C", 1200, "Dell", "CC-99"]])
items, esito = rows_from_workbook(p, STANZE)
assert len(items) == 1 and esito["scartate"] == 0
assert items[0]["modello"] == "T14 Gen 4" and items[0]["seriale"] == "PF4A1B2C"
assert esito["colonne_ignorate"] == ["Costo", "Fornitore", "Centro di costo"], esito
avvisi = ImportDialog._avvertenze(esito)
assert any("Costo" in a and "non verra' importato" in a for a in avvisi), avvisi

# ---- piu' di sei colonne ignote: l'elenco si accorcia
p = foglio("molte", ["Asset Tag", "Modello"] + ["Extra%d" % i for i in range(9)],
           [["IT-0102", "T14"] + [""] * 9])
_, esito = rows_from_workbook(p, STANZE)
assert len(esito["colonne_ignorate"]) == 9
avviso = ImportDialog._avvertenze(esito)[0]
assert "e altre 3" in avviso, avviso

# ---- manca il modello: si importa, ma il programma lo dice
p = foglio("senza_modello", ["Asset Tag", "Tipo", "Numero di serie"],
           [["IT-0103", "Laptop", "PF5NEW1"], ["IT-0104", "Laptop", "PF5NEW2"]])
items, esito = rows_from_workbook(p, STANZE)
assert len(items) == 2 and esito["senza_modello"] == 2
assert all(i["modello"] == "" for i in items)
assert any("non hanno il modello" in a for a in ImportDialog._avvertenze(esito))
s = InventoryStore(os.path.join(d, "Inv.xlsx"), iphone_room=fixture.BAU)
s.create_if_missing()
assert s.import_items(items, "merge")["aggiunti"] == 2, "si importano lo stesso"

# ---- nessuna colonna identificativa: errore chiaro, niente importato
p = foglio("inutile", ["Costo", "Fornitore"], [[100, "Dell"]])
try:
    rows_from_workbook(p, STANZE); raise SystemExit("file senza identificativo accettato")
except InventoryError as e:
    assert "Asset Tag" in str(e) and "IMEI" in str(e), str(e)

# ---- un file con titolo e riga vuota in cima: le intestazioni si trovano lo stesso
p = foglio("con_titolo", ["Digital Kiosk", None, None, None],
           [["Esportato il 30/08/2026", None, None, None],
            [None, None, None, None],
            ["Asset Tag", "Tipo", "Modello", "Numero di serie"],
            ["IT-0900", "Laptop", "T14 Gen 5", "PF5NEW9"]])
items, esito = rows_from_workbook(p, STANZE)
assert len(items) == 1 and items[0]["asset_tag"] == "IT-0900", items
assert items[0]["modello"] == "T14 Gen 5"

# ---- ma un titolo lunghissimo senza tabella resta un errore
p = foglio("solo_titoli", ["Relazione annuale"],
           [["Reparto"], ["Nota"], ["Altro"], ["Ancora"], ["E ancora"],
            ["Sesto"], ["Settimo"], ["Ottavo"], ["Nono"], ["Decimo"],
            ["Undicesimo"], ["Dodicesimo"],
            ["Asset Tag"], ["IT-0901"]])
try:
    rows_from_workbook(p, STANZE); raise SystemExit("intestazione troppo in basso accettata")
except InventoryError as e:
    assert "intestazioni" in str(e)

# ---- maiuscole, spazi e sinonimi
intestazioni = ["  ASSET TAG ", "tipo", "MoDeLLo", "s/n", "ubicazione"]
mappa = map_headers(intestazioni)
assert set(mappa.values()) == {"asset_tag", "tipo", "modello", "seriale", "stanza"}, mappa

# ---- due colonne per lo stesso campo: vince la prima, la seconda e' segnalata
p = foglio("doppie", ["Asset Tag", "Tag", "Modello", "Descrizione"],
           [["IT-0105", "ALTRO", "Vero modello", "Altra descrizione"]])
items, esito = rows_from_workbook(p, STANZE)
assert items[0]["asset_tag"] == "IT-0105" and items[0]["modello"] == "Vero modello"
assert esito["colonne_ignorate"] == ["Tag", "Descrizione"], esito

# ---- un file corretto non genera avvertenze
p = foglio("pulito", ["Asset Tag", "Tipo", "Modello", "Numero di serie"],
           [["IT-0106", "Laptop", "T14 Gen 5", "PF5K9M8F"]])
_, esito = rows_from_workbook(p, STANZE)
assert esito["colonne_ignorate"] == [] and esito["senza_modello"] == 0
assert ImportDialog._avvertenze(esito) == []
print("IMPORT ROBUSTO OK")

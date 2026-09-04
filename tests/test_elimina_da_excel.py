"""Eliminare in blocco partendo da un file Excel, e il cestino che si riempie.

Chi deve dismettere ottanta portatili ha un foglio, non una colonna da
incollare. Il pulsante rosso in home legge il file, dice che cosa sparisce
stanza per stanza, e per confermare chiede di scrivere la parola.

Il cestino tiene un numero fisso di record. Quando quelli che si eliminano non
ci stanno tutti, quello che esce e' perso per sempre: non si decide da soli, si
chiede - o entrano al posto dei piu' vecchi, o si cancellano subito quelli in
eccesso.
"""
import os, sys, tempfile
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import fixture
from openpyxl import Workbook
from tkinter import filedialog, messagebox
from inventario import ui
from inventario.store import (ELIMINATI_MASSIMO, MASSIMO_ELIMINA_EXCEL,
                              InventoryStore, new_item, righe_da_workbook)
from inventario.ui import App, CestinoPienoDialog, EliminaDaExcelDialog

BAU, KIOSK, DR = fixture.BAU, fixture.KIOSK, fixture.DR
INTESTAZIONI = ["Asset Tag", "Tipo", "Modello", "Numero di serie", "Stato", "Note"]

def scrivi(righe, titolo="Inventario"):
    wb = Workbook(); ws = wb.active; ws.title = titolo
    for r in righe:
        ws.append(r)
    p = os.path.join(tempfile.mkdtemp(), "elenco.xlsx")
    wb.save(p); wb.close()
    return p

# ===================== leggere il file, in tutte le forme =====================

# ---- un'esportazione del programma: si prendono solo le colonne che contano
completo = scrivi([INTESTAZIONI,
                   ["SITE SERVICES BAU", None, None, None, None, None],
                   ["IT-0101", "Laptop", "T14", "PF1", "Disponibile", "nota"],
                   ["IT-0102", "Laptop", "T14", "PF2", "Disponibile", None]])
righe, letto = righe_da_workbook(completo)
assert righe == ["SITE SERVICES BAU", "IT-0101", "IT-0102"], righe
assert "T14" not in "".join(righe), "delle altre colonne non importa niente"

# ---- una colonna sola, incollata a mano e salvata senza intestazioni
nuda = scrivi([["IT-0101"], ["IT-0102"], [None], ["IT-0103"]])
righe, letto = righe_da_workbook(nuda)
assert righe == ["IT-0101", "IT-0102", "IT-0103"], righe
assert letto["vuote"] == 1, letto

# ---- un file senza niente dentro
vuoto = scrivi([[None, None], [None, None]])
righe, _ = righe_da_workbook(vuoto)
assert righe == [], righe

# ===================== la regola del cestino pieno =====================
store = InventoryStore(fixture.build(), iphone_room=BAU)
store.load()
assert store.eccesso_cestino(3) == 0, "il cestino e' vuoto: ci stanno tutti"
assert store.eccesso_cestino(ELIMINATI_MASSIMO + 5) == 5, store.eccesso_cestino(1)

# riempiamo il cestino fin quasi all'orlo
tanti = [new_item("IT-9%03d" % n, "Laptop", "T14", "PF%d" % n, KIOSK)
         for n in range(ELIMINATI_MASSIMO - 2)]
for item in tanti:
    store.add(item)
store.load()
store.delete([i["asset_tag"] for i in tanti])
store.load()
assert len(store.eliminati()) == ELIMINATI_MASSIMO - 2, len(store.eliminati())

# ---- tre in piu': uno non ci sta
assert store.eccesso_cestino(3) == 1, store.eccesso_cestino(3)

# ---- "al posto dei piu' vecchi": entrano tutti, il cestino resta al massimo
prima_vecchi = [v["asset_tag"] for v in store.eliminati()][-3:]
store.delete(["IT-0101", "IT-0102", "IT-0103"], in_eccesso="cestino")
store.load()
dentro = [v["asset_tag"] for v in store.eliminati()]
assert len(dentro) == ELIMINATI_MASSIMO, len(dentro)
assert "IT-0103" in dentro and "IT-0101" in dentro, "i nuovi ci sono tutti"
assert prima_vecchi[-1] not in dentro, "il piu' vecchio e' uscito"
assert store.cancellati_per_sempre == [], "nessuno e' stato cancellato al volo"

# ---- "cancella quelli in eccesso": il cestino non cresce, e lo dice
store = InventoryStore(fixture.build(), iphone_room=BAU)
store.load()
for item in tanti:
    store.add(item)
store.load()
store.delete([i["asset_tag"] for i in tanti])
store.load()
store.delete(["IT-0101", "IT-0102", "IT-0103"], in_eccesso="definitivo")
store.load()
dentro = [v["asset_tag"] for v in store.eliminati()]
assert len(dentro) == ELIMINATI_MASSIMO, len(dentro)
assert [i["asset_tag"] for i in store.cancellati_per_sempre] == ["IT-0103"], \
    store.cancellati_per_sempre
assert "IT-0103" not in dentro, "quello in eccesso non e' nel cestino"
assert "IT-0101" in dentro and "IT-0102" in dentro, "gli altri due ci stavano"
assert not [i for i in store.items if i["asset_tag"] == "IT-0103"], \
    "dall'inventario e' sparito comunque"

# ---- e in tutti e due i casi l'inventario e' quello giusto
for tag in ("IT-0101", "IT-0102", "IT-0103"):
    assert not [i for i in store.items if i["asset_tag"] == tag], tag

# ===================== le finestre =====================
app = App(fixture.build())
app._initial_load()
app.update()
avvisi = []
messagebox.showinfo = lambda t, m, **k: avvisi.append((t, m))
messagebox.showwarning = lambda t, m, **k: avvisi.append((t, m))
messagebox.showerror = lambda t, m, **k: avvisi.append((t, m))

# ---- il cestino pieno: le due risposte e l'annullamento
d = CestinoPienoDialog(app, 4, 10, ELIMINATI_MASSIMO - 6, ELIMINATI_MASSIMO)
d.var_come.set("definitivo"); d._ok()
assert d.result == "definitivo", d.result
d = CestinoPienoDialog(app, 4, 10, ELIMINATI_MASSIMO - 6, ELIMINATI_MASSIMO)
d._ok()
assert d.result == "cestino", "la risposta prudente e' quella predefinita"
d = CestinoPienoDialog(app, 4, 10, ELIMINATI_MASSIMO - 6, ELIMINATI_MASSIMO)
d._cancel()
assert d.result is None

# ---- finche' ci stanno tutti non si chiede niente
chiesto = []
class NonSiApre:
    def __init__(self, *a, **k): chiesto.append(a)
    def show(self): return "cestino"
ui.CestinoPienoDialog = NonSiApre
assert app._regola_cestino(2) == "cestino"
assert not chiesto, "il cestino era vuoto: non c'era niente da chiedere"

# ---- la finestra dell'eliminazione da Excel: legge, riepiloga, chiede la parola
elenco = scrivi([["IT-0101"], ["IT-0107"], ["IT-9999"]])
filedialog.askopenfilename = lambda **k: elenco
d = EliminaDaExcelDialog(app, app.store)
d._scegli()
testo = d.riepilogo.get("1.0", "end")
assert "VERRANNO ELIMINATI: 1" in testo, testo
assert BAU in testo and "IT-0101" in testo, "dice da che stanza sparisce"
assert "non sono in inventario: 1" in testo and "IT-9999" in testo, testo
assert "non si possono eliminare: 1" in testo and "IT-0107" in testo, \
    "il portatile in prestito e' bloccato, e si dice perche'"
assert str(d.btn_elimina.state()) .find("disabled") < 0, "si puo' eliminare"

# ---- senza la parola non si elimina
avvisi.clear()
d.var_conferma.set("si"); d._ok()
assert d.result is None and avvisi and avvisi[-1][0] == "Conferma non valida", avvisi
d.var_conferma.set(ui.PAROLA_ELIMINA); d._ok()
assert d.result == ["IT-0101"], d.result

# ---- un file senza codici buoni: niente da confermare
filedialog.askopenfilename = lambda **k: scrivi([["ZZ-1"], ["ZZ-2"]])
d = EliminaDaExcelDialog(app, app.store)
d._scegli()
assert "Non c'e' niente da eliminare." in d.riepilogo.get("1.0", "end")
assert "disabled" in str(d.btn_elimina.state()), "non si elimina niente"

# ---- oltre il massimo il file si rifiuta, e si dice quanti erano
grande = App(fixture.build())
grande._initial_load()
molti = [new_item("IT-7%03d" % n, "Laptop", "T14", "PF%d" % n, KIOSK)
         for n in range(MASSIMO_ELIMINA_EXCEL + 1)]
for item in molti:
    grande.store.add(item)
grande.store.load()
filedialog.askopenfilename = lambda **k: scrivi([[i["asset_tag"]] for i in molti])
d = EliminaDaExcelDialog(grande, grande.store)
d._scegli()
testo = d.riepilogo.get("1.0", "end")
assert "%d dispositivi in inventario" % (MASSIMO_ELIMINA_EXCEL + 1) in testo, testo
assert "al massimo %d per volta" % MASSIMO_ELIMINA_EXCEL in testo, testo
assert "disabled" in str(d.btn_elimina.state()), "non si elimina niente"
grande.destroy()

# ===================== il giro completo dalla home =====================
prima = len(app.store.items)
ui.EliminaDaExcelDialog = lambda parent, store: type(
    "Finta", (), {"show": lambda self: ["IT-0102", "IT-0103"]})()
ui.CestinoPienoDialog = NonSiApre
avvisi.clear()
app.on_elimina_da_excel()
app.store.load()
assert len(app.store.items) == prima - 2, len(app.store.items)
titolo, corpo = avvisi[-1]
assert titolo == "Eliminazione completata", avvisi[-1]
assert "2 dispositivi eliminati." in corpo, corpo
assert "Copia di sicurezza" in corpo, corpo
assert sorted(v["asset_tag"] for v in app.store.eliminati()) == ["IT-0102", "IT-0103"], \
    "sono finiti negli eliminati di recente"

# ---- annullando la domanda sul cestino non si elimina niente
prima = len(app.store.items)
ui.CestinoPienoDialog = lambda *a, **k: type("Finta", (), {"show": lambda self: None})()
app._regola_cestino = lambda quanti: None
app.on_elimina_da_excel()
app.store.load()
assert len(app.store.items) == prima, "annullato vuol dire annullato"

app.destroy()
print("ELIMINA DA EXCEL OK")

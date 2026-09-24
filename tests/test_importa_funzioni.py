"""Impostazioni > Importa Asset Function da foglio Excel.

Quello che e' stato chiesto:
- legge la colonna Asset Tag e la colonna Asset Function - o comunque
  qualsiasi colonna che contiene SOLO Standard e PC Refresh;
- aggiorna l'inventario vivo con il valore di ogni asset tag;
- le stanze dei dispositivi presenti non cambiano, nessuno si sposta;
- i record del foglio che nell'inventario non ci sono, si aggiungono.
"""
import os, sys, tempfile, time
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import fixture
from openpyxl import Workbook
from tkinter import filedialog, messagebox
from inventario import ui
from inventario.store import InventoryStore, funzioni_da_workbook, new_item
from inventario.ui import App

BAU, KIOSK, DR = fixture.BAU, fixture.KIOSK, fixture.DR
STANZE = [BAU, KIOSK, DR]

def foglio(righe):
    wb = Workbook(); ws = wb.active; ws.title = "Foglio"
    for r in righe:
        ws.append(r)
    p = os.path.join(tempfile.mkdtemp(), "funzioni.xlsx"); wb.save(p); wb.close()
    return p

# ======================= leggere la colonna =======================
# ---- per nome
p = foglio([["Asset Tag", "Asset Function"],
            ["IT-0101", "PC Refresh"], ["IT-0102", "standard"]])
items, esito = funzioni_da_workbook(p, STANZE)
assert esito["colonna_funzione"] == "Asset Function", esito["colonna_funzione"]
assert dict((i["asset_tag"], i["funzione"]) for i in items) == \
    {"IT-0101": "PC Refresh", "IT-0102": "Standard"}, items

# ---- per contenuto: la colonna si chiama come vuole, ma contiene solo quei due
p = foglio([["Asset Tag", "Reparto", "Programma 2026", "Note"],
            ["IT-0101", "Acquisti", "PC Refresh", "da ritirare"],
            ["IT-0102", "Vendite", "Standard", ""],
            ["IT-0103", "Standard", "Standard", "nota"],      # "Reparto" non e' solo funzioni
            ["IT-0104", "HR", "", "manca"]])
items, esito = funzioni_da_workbook(p, STANZE)
assert esito["colonna_funzione"] == "Programma 2026", esito["colonna_funzione"]
funz = dict((i["asset_tag"], i["funzione"]) for i in items)
assert funz == {"IT-0101": "PC Refresh", "IT-0102": "Standard",
                "IT-0103": "Standard", "IT-0104": ""}, funz

# ---- nessuna colonna utile: lo si dice
p = foglio([["Asset Tag", "Note"], ["IT-0101", "boh"]])
_i, esito = funzioni_da_workbook(p, STANZE)
assert esito["colonna_funzione"] is None

# ======================= l'archivio: solo la funzione =======================
store = InventoryStore(fixture.build(), iphone_room=BAU)
store.stanze = STANZE
store.load()
prima = dict((i["asset_tag"], dict(i)) for i in store.items)
a = store.anteprima_funzioni({"IT-0101": "PC Refresh", "IT-0107": "Standard",
                              "IT-9999": "Standard"})
assert [c[0] for c in a["cambi"]] == ["IT-0101", "IT-0107"], a
assert a["assenti"] == ["IT-9999"], a
cambi = store.aggiorna_funzioni({"IT-0101": "PC Refresh", "IT-0107": "Standard",
                                 "IT-9999": "Standard"})
assert sorted(c[0] for c in cambi) == ["IT-0101", "IT-0107"], cambi
store.load()
dopo = dict((i["asset_tag"], i) for i in store.items)
assert dopo["IT-0101"]["funzione"] == "PC Refresh"
assert dopo["IT-0107"]["funzione"] == "Standard", "anche a un dispositivo in prestito"
# niente altro si e' mosso: stanza, stato, prestito, note, modello, seriale
for tag in ("IT-0101", "IT-0107"):
    for campo in ("stanza", "stato", "prestato_a", "note", "modello", "seriale", "tipo"):
        assert dopo[tag][campo] == prima[tag][campo], (tag, campo)
assert "IT-9999" not in dopo, "l'archivio non aggiunge: lo fa chi chiama"
# un valore non valido non cancella quello che c'e'
store.aggiorna_funzioni({"IT-0101": "Ricondizionato"})
store.load()
assert [i for i in store.items if i["asset_tag"] == "IT-0101"][0]["funzione"] == "PC Refresh"

# ======================= dalle Impostazioni, il giro intero =======================
app = App(fixture.build())
app._initial_load()
app.update()
scadenza = time.time() + 0.4
while time.time() < scadenza:
    app.update()
avvisi = []
for nome in ("showinfo", "showwarning", "showerror"):
    setattr(messagebox, nome, lambda t, m, **k: avvisi.append((t, m)))
riepiloghi = []
class Conferma:
    def __init__(self, parent, titolo, testa, righe, *a, **k):
        riepiloghi.append((testa, righe))
    def show(self): return True
ui.ConfermaOperazioneDialog = Conferma
VeraStanze = ui.StanzeDaAssegnareDialog

prima = dict((i["asset_tag"], dict(i)) for i in app.store.items)
# il foglio: due presenti, uno nuovo con la sua stanza, uno nuovo senza stanza
p = foglio([["Asset Tag", "Tipo", "Stanza", "Funzione"],
            ["IT-0101", "Laptop", DR, "PC Refresh"],        # e' in BAU: NON si sposta
            ["IT-0106", "Laptop", BAU, "Standard"],         # e' nel Kiosk: NON si sposta
            ["IT-8801", "Laptop", KIOSK, "PC Refresh"],     # nuovo, con stanza
            ["IT-8802", "Laptop", "", "Standard"]])         # nuovo, stanza da chiedere
filedialog.askopenfilename = lambda **k: p
chieste = []
class UnaStanza:
    def __init__(self, parent, quanti, stanze, nuovi=False, codici=()):
        chieste.append((quanti, nuovi, list(codici)))
    def show(self): return {"come": "una", "stanza": DR}
ui.StanzeDaAssegnareDialog = UnaStanza

app.on_importa_funzioni()
app.store.load()
dopo = dict((i["asset_tag"], i) for i in app.store.items)
# ---- i presenti: funzione aggiornata, stanza intatta
assert dopo["IT-0101"]["funzione"] == "PC Refresh"
assert dopo["IT-0101"]["stanza"] == prima["IT-0101"]["stanza"] == BAU, "non si sposta"
assert dopo["IT-0106"]["funzione"] == "Standard"
assert dopo["IT-0106"]["stanza"] == prima["IT-0106"]["stanza"] == KIOSK, "non si sposta"
# ---- i nuovi: aggiunti, nella stanza del foglio o in quella chiesta
assert dopo["IT-8801"]["stanza"] == KIOSK and dopo["IT-8801"]["funzione"] == "PC Refresh"
assert dopo["IT-8802"]["stanza"] == DR and dopo["IT-8802"]["funzione"] == "Standard"
assert chieste == [(1, True, ["IT-8802"])], \
    "si chiede solo per chi non ce l'ha, dicendo che e' nuovo e quale: %s" % chieste
# ---- il riepilogo prima di scrivere lo diceva
testa, righe = riepiloghi[-1]
testo = "\n".join(righe)
assert "Aggiornare 2 Asset Function e aggiungere 2 dispositivi?" in testa, testa
assert "DA AGGIORNARE: 2" in testo and "DA AGGIUNGERE, non sono in inventario: 2" in testo, testo
assert "Funzione" in testo, "dice da quale colonna ha letto"
# ---- e alla fine lo ridice, con la copia di sicurezza
assert avvisi[-1][0] == "Asset Function importate", avvisi[-1]
assert "Asset Function aggiornate: 2" in avvisi[-1][1] and "Dispositivi aggiunti: 2" in avvisi[-1][1]
assert "Copia di sicurezza" in avvisi[-1][1]

# ---- rifatto uguale: niente da fare, e lo dice
avvisi.clear()
app.on_importa_funzioni()
assert avvisi[-1][0] == "Niente da aggiornare", avvisi[-1]

# ---- un foglio senza la colonna: non tocca niente
avvisi.clear()
filedialog.askopenfilename = lambda **k: foglio([["Asset Tag", "Note"], ["IT-0101", "x"]])
app.on_importa_funzioni()
assert avvisi[-1][0] == "Nessuna Asset Function nel foglio", avvisi[-1]

# ---- la finestra vera dice che sono nuovi, e quali
ui.StanzeDaAssegnareDialog = VeraStanze
d = ui.StanzeDaAssegnareDialog(app, 2, STANZE, nuovi=True, codici=["IT-7001", "IT-7002"])
testi_d = []
def raccogli(w):
    for c in w.winfo_children():
        try: testi_d.append(str(c.cget("text")))
        except Exception: pass
        raccogli(c)
raccogli(d)
assert "2 dispositivi nuovi, senza stanza" in testi_d, testi_d
assert any("Non sono ancora in inventario: IT-7001, IT-7002" in t for t in testi_d), testi_d
assert any("In che stanza li aggiungo?" in t for t in testi_d), testi_d
d.destroy()
d = ui.StanzaDelDispositivoDialog(app, new_item("IT-7001", "Laptop"), STANZE, 1, 2, nuovo=True)
testi_d = []; raccogli(d)
assert "Dispositivo nuovo 1 di 2, non ancora in inventario" in testi_d, testi_d
d.destroy()

# ---- il pulsante sta nelle Impostazioni
d = ui.RoomsDialog(app, app.cfg["rooms"], app.cfg["types"],
                   app.cfg.get("loan_rooms", []), app.cfg.get("iphone_room", ""))
testi = []
def gira(w):
    for c in w.winfo_children():
        try: testi.append(str(c.cget("text")))
        except Exception: pass
        gira(c)
gira(d)
assert "Importa Asset Function da foglio Excel..." in testi, testi
d._importa_funzioni()
assert d.result == {"importa_funzioni": True}

app.destroy()
print("IMPORTA FUNZIONI OK")

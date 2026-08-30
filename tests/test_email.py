"""L'invio per e-mail offerto in coda a ogni esportazione."""
import os, sys, tempfile, zipfile
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import fixture
from tkinter import messagebox
from inventario import excel_io
from inventario.store import InventoryError
from inventario.ui import App, EsportazioneFattaDialog

BAU, KIOSK, DR = fixture.BAU, fixture.KIOSK, fixture.DR
d = tempfile.mkdtemp()
avvisi = []
messagebox.showwarning = lambda t, m, **k: avvisi.append((t, m))
app = App(fixture.build()); app._initial_load()

# ---- la finestra di coda offre le tre scelte, non ne aggiunge ai menu
uno = os.path.join(d, "Inventario_20260831.xlsx")
excel_io.export(app.store.items, uno)
dlg = EsportazioneFattaDialog(app, "13 dispositivi esportati", [uno])
corpo = dlg.winfo_children()[0]
pulsanti = [b.cget("text") for b in corpo.winfo_children() if b.winfo_class() == "TButton"]
assert pulsanti == ["Invia per e-mail con Outlook", "Apri il file", "Ho finito"], pulsanti
dlg._scegli("email")
assert dlg.result == "email"

dlg = EsportazioneFattaDialog(app, "x", [uno]); dlg._cancel()
assert dlg.result is None, "chiudendo non si fa niente"

# ---- l'elenco dei file si accorcia quando sono tanti
molti = []
for n in range(9):
    p = os.path.join(d, "Stanza_%d.xlsx" % n)
    excel_io.export(app.store.items[:2], p)
    molti.append(p)
dlg = EsportazioneFattaDialog(app, "9 file", molti)
etichette = [w.cget("text") for w in dlg.winfo_children()[0].winfo_children()
             if w.winfo_class() == "TLabel"]
assert any("e altri 3" in t for t in etichette), etichette
dlg._cancel()

# ---- senza Outlook il file resta comunque prodotto, e lo si dice
assert excel_io.outlook_disponibile() is False, "su questo Mac non c'e' Outlook"
try:
    excel_io.allega_a_outlook([uno]); raise SystemExit("allegato senza Outlook")
except InventoryError as e:
    assert "Outlook non e' stato trovato" in str(e)
    assert "e' stato creato lo stesso" in str(e), "va detto che il file c'e'"
assert os.path.exists(uno), "il file non deve sparire"

# la stessa cosa passando dall'interfaccia: avviso, e nessuna eccezione
avvisi.clear()
app.fine_esportazione = app.fine_esportazione   # niente monkeypatch: si usa quella vera
class FintaScelta(EsportazioneFattaDialog):
    def show(self):
        return "email"
import inventario.ui as ui
ui.EsportazioneFattaDialog = FintaScelta
app.fine_esportazione("13 dispositivi", [uno])
assert avvisi and avvisi[-1][0] == "Invio per e-mail", avvisi
assert "Outlook non e' stato trovato" in avvisi[-1][1]

# ---- niente da allegare
try:
    excel_io.allega_a_outlook([]); raise SystemExit("elenco vuoto accettato")
except InventoryError as e:
    assert "nessun file da allegare" in str(e)

# ---- con piu' file si prepara un archivio unico
# si simulano Outlook e il suo avvio: qui non c'e' e non deve esserci
lanci = []
excel_io._percorso_outlook = lambda: os.path.join(d, "OUTLOOK.EXE")
excel_io.subprocess.Popen = lambda comando, *a, **k: lanci.append(comando)
allegato = excel_io.allega_a_outlook(molti)
assert allegato.endswith(".zip"), allegato
with zipfile.ZipFile(allegato) as archivio:
    dentro = archivio.namelist()
assert len(dentro) == 9 and all(n.endswith(".xlsx") for n in dentro), dentro
assert all(os.path.basename(p) in dentro for p in molti)

# con un file solo si allega quello, senza archivio
allegato = excel_io.allega_a_outlook([uno])
assert allegato == uno, allegato
# il comando passato a Outlook e' quello documentato: /a seguito dall'allegato
assert len(lanci) == 2, lanci
assert lanci[-1][1] == "/a" and lanci[-1][2] == uno, lanci[-1]
assert lanci[0][2].endswith(".zip")
app.destroy()
print("EMAIL OK")

"""Ripartire da una copia locale quando la cartella di rete non c'e' piu'.

I backup automatici stanno sulla share, accanto ai dati: coprono l'errore
umano, non la share che sparisce. La copia locale invece se la porta via il
tecnico, e da sola deve bastare a rimettere tutto - non solo i dispositivi, ma
anche le stanze, che nel file dei dispositivi non ci sono.
"""
import json, os, sys, tempfile, zipfile
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import fixture
from tkinter import filedialog, messagebox
from inventario import config, ui
from inventario.store import (NOME_DATI_NELLO_ZIP, NOME_ELIMINATI_NELLO_ZIP,
                              NOME_IMPOSTAZIONI_NELLO_ZIP,
                              InventoryError, InventoryStore, new_item)
from inventario.ui import App, riepilogo_copia_locale

BAU, KIOSK, DR = fixture.BAU, fixture.KIOSK, fixture.DR
fuori = tempfile.mkdtemp()

percorso = fixture.build()
store = InventoryStore(percorso, iphone_room=BAU)
store.load()

# ---- la copia locale e' uno zip con dentro i dispositivi e le impostazioni
archivio = os.path.join(fuori, "Copia.zip")
salvato, impostazioni, quanti = store.copia_in(archivio)
assert salvato == archivio and quanti == 13, (salvato, quanti)
assert impostazioni == NOME_IMPOSTAZIONI_NELLO_ZIP, impostazioni
dentro = zipfile.ZipFile(archivio).namelist()
# il cestino viaggia sempre, anche vuoto: senza, il ripristino non saprebbe che
# in quel momento era vuoto e lascerebbe dentro quello di adesso
assert dentro == [NOME_DATI_NELLO_ZIP, NOME_IMPOSTAZIONI_NELLO_ZIP,
                  NOME_ELIMINATI_NELLO_ZIP], dentro

# ---- l'anteprima dice che cosa tornerebbe, senza toccare niente
rapporto = store.anteprima_copia_locale(archivio)
assert rapporto["dispositivi"] == 13, rapporto
assert rapporto["per_stanza"] == {BAU: 5, KIOSK: 5, DR: 3}, rapporto["per_stanza"]
assert rapporto["impostazioni"]["rooms"] == [BAU, KIOSK, DR], rapporto["impostazioni"]
assert rapporto["quando"] is not None
assert len(store.load()) == 13, "l'anteprima non deve toccare l'inventario"

# ---- il disastro: si cambiano le stanze e si svuota mezzo inventario
config.save_shared_config(percorso, {"rooms": ["Stanza sbagliata"],
                                     "types": ["Laptop"],
                                     "loan_rooms": [],
                                     "iphone_room": "Stanza sbagliata"})
store.delete(["IT-0101", "IT-0102", "IT-0103"])
store.load()
assert len(store.items) == 10
assert config.load_shared_config(percorso)["rooms"] == ["Stanza sbagliata"]

# ---- e si torna indietro da un file che sta fuori dalla rete
tornati, precedente, con_impostazioni = store.ripristina_da_copia_locale(archivio)
assert tornati == 13, tornati
assert con_impostazioni is True
assert os.path.exists(precedente), "lo stato sbagliato resta recuperabile"
store.load()
assert len(store.items) == 13
assert [i["asset_tag"] for i in store.items if i["asset_tag"] == "IT-0101"]
tornate = config.load_shared_config(percorso)
assert tornate["rooms"] == [BAU, KIOSK, DR], tornate
assert tornate["loan_rooms"] == [KIOSK], tornate
assert tornate["iphone_room"] == BAU, tornate
# e i prestiti in corso sono ancora li': e' l'inventario, non un estratto
assert sorted(i["asset_tag"] for i in store.items if i.get("prestato_a")) == \
    ["IT-0107", "IT-0110"]

# ---- una copia vecchia, il solo .xlsx, resta ripristinabile
vecchia = os.path.join(fuori, "Vecchia.xlsx")
store.copia_in(vecchia)
assert os.path.exists(os.path.splitext(vecchia)[0] + "_impostazioni.json")
store.delete(["IT-0101"])
tornati, _precedente, con_impostazioni = store.ripristina_da_copia_locale(vecchia)
assert tornati == 13 and con_impostazioni is True, (tornati, con_impostazioni)

# ---- un .xlsx senza il suo file di impostazioni: tornano solo i dispositivi
solo_dati = os.path.join(fuori, "SoloDati.xlsx")
store.copia_in(solo_dati)
os.remove(os.path.splitext(solo_dati)[0] + "_impostazioni.json")
config.save_shared_config(percorso, {"rooms": ["Solo questa"], "types": ["Laptop"],
                                     "loan_rooms": [], "iphone_room": "Solo questa"})
tornati, _p, con_impostazioni = store.ripristina_da_copia_locale(solo_dati)
assert tornati == 13 and con_impostazioni is False
assert config.load_shared_config(percorso)["rooms"] == ["Solo questa"], \
    "senza impostazioni nella copia, quelle di adesso non si toccano"
store.ripristina_da_copia_locale(archivio)          # si rimette a posto

# ---- un file rovinato viene rifiutato senza toccare niente
rotto = os.path.join(fuori, "Rotto.zip")
open(rotto, "wb").write(b"questo non e' uno zip")
prima = len(store.load())
try:
    store.ripristina_da_copia_locale(rotto)
    raise SystemExit("ha ripristinato da un file rovinato")
except InventoryError as exc:
    assert "non e' una copia leggibile" in str(exc), str(exc)
assert len(store.load()) == prima, "l'inventario non doveva essere toccato"

# ---- uno zip senza inventario dentro: stesso trattamento
vuoto = os.path.join(fuori, "Vuoto.zip")
with zipfile.ZipFile(vuoto, "w") as z:
    z.writestr("lettera.txt", "niente inventario qui")
try:
    store.ripristina_da_copia_locale(vuoto)
    raise SystemExit("ha ripristinato da uno zip senza inventario")
except InventoryError as exc:
    assert "nessun inventario" in str(exc), str(exc)
assert len(store.load()) == prima

# ---- il riepilogo dice le stanze prima e dopo, e che cosa cambia
rapporto = store.anteprima_copia_locale(archivio)
testo = "\n".join(riepilogo_copia_locale(
    rapporto, 10, {BAU: 2, "Stanza sbagliata": 8},
    {"rooms": ["Stanza sbagliata"], "types": ["Laptop"], "loan_rooms": [],
     "iphone_room": "Stanza sbagliata"}))
assert "La copia contiene 13 dispositivi" in testo, testo
assert "COME RESTANO LE STANZE:" in testo, testo
assert "2  ->  5" in testo, testo                 # BAU torna a cinque
assert "8  ->  0" in testo, testo                 # la stanza inventata si svuota
assert "TORNANO ANCHE LE IMPOSTAZIONI:" in testo, testo
assert "* Stanze:" in testo, "le righe che cambiano vanno marcate"

# ---- e se la copia non le porta, lo dice
senza = dict(rapporto); senza["impostazioni"] = None
testo = "\n".join(riepilogo_copia_locale(senza, 10, {}, {}))
assert "non porta le impostazioni" in testo, testo

# ============================ dall'interfaccia ============================
app = App(fixture.build())
app._initial_load()
app.update()
avvisi = []
messagebox.showinfo = lambda t, m, **k: avvisi.append((t, m))
messagebox.showerror = lambda t, m, **k: avvisi.append((t, m))
ui.ConfermaOperazioneDialog.show = lambda self: True

# si salva una copia, si fa un disastro, si torna indietro dal menu
locale = os.path.join(fuori, "DallaApp.zip")
filedialog.asksaveasfilename = lambda **k: locale
app.on_copia_locale()
assert os.path.exists(locale), avvisi
assert "Ripristina da un file locale" in avvisi[-1][1], avvisi[-1][1]

app._run(lambda: app.store.delete(["IT-0101", "IT-0102"]), "ok")
config.save_shared_config(app.store.path, {"rooms": ["Rovinata"], "types": ["Laptop"],
                                           "loan_rooms": [], "iphone_room": "Rovinata"})
app.cfg = config.load_shared_config(app.store.path)
filedialog.askopenfilename = lambda **k: locale
app.on_ripristino_locale()
app.store.load()
assert len(app.store.items) == 13, len(app.store.items)
assert app.cfg["rooms"] == [BAU, KIOSK, DR], app.cfg["rooms"]
titolo, corpo = avvisi[-1]
assert titolo == "Inventario ripristinato", avvisi[-1]
assert "Sono tornate anche le impostazioni" in corpo, corpo
assert "Lo stato precedente e' stato salvato" in corpo, corpo

# ---- annullando la scelta del file non succede niente
quanti_prima = len(app.store.items)
filedialog.askopenfilename = lambda **k: ""
app.on_ripristino_locale()
assert len(app.store.load()) == quanti_prima

# ---- e un file rovinato lo dice senza toccare l'inventario
filedialog.askopenfilename = lambda **k: rotto
app.on_ripristino_locale()
assert "non e' una copia leggibile" in avvisi[-1][1], avvisi[-1][1]
assert len(app.store.load()) == quanti_prima

app.destroy()
print("RIPRISTINO LOCALE OK")

"""Un identificativo non puo' stare in due posti, e non si riscrive di nascosto.

L'identificativo, e solo quello, decide: l'unico campo che non puo' mai
ripetersi e' l'**asset tag** - per gli iPhone l'IMEI. Tutto il resto della
scheda puo' variare quanto vuole, e nella decisione non viene nemmeno guardato:
due schede diverse in ogni campo con lo stesso asset tag sono lo stesso
dispositivo, due schede identiche in tutto con asset tag diversi sono due
dispositivi.

Tre regole, che valgono da ogni strada che rimette dentro un dispositivo -
importazione da qualsiasi pagina, inserimento singolo, controllo doppioni:

1. gia' in inventario  -> non si importa MAI: la riga si salta e si dice in che
   stanza sta quello che c'e' gia'. La sua scheda non viene riscritta da un
   foglio;
2. solo nel cestino    -> si importa, e la voce nel cestino sparisce: il
   dispositivo e' tornato in inventario e non puo' restare anche fra gli
   eliminati;
3. da nessuna parte    -> entra normalmente.
"""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import fixture
from tkinter import messagebox
from inventario import ui
from inventario.store import InventoryError, InventoryStore, new_item
from inventario.ui import App

BAU, KIOSK, DR = fixture.BAU, fixture.KIOSK, fixture.DR

def prepara():
    """Inventario con IT-0101 dentro, IT-0106 nel cestino, IT-9001 inesistente."""
    store = InventoryStore(fixture.build(), iphone_room=BAU)
    store.load()
    store.delete(["IT-0106"])
    return store

# ============================ importazione ============================
store = prepara()
prima = dict([i for i in store.items if i["asset_tag"] == "IT-0101"][0])
# la riga di IT-0101 ha lo stesso asset tag e TUTTO il resto diverso - tipo,
# modello, seriale, stanza, note: deve bastare il tag a riconoscerla
righe = [new_item("IT-0101", "Tablet", "RISCRITTO DAL FOGLIO", "XXX", KIOSK, "nota nuova"),
         new_item("IT-0106", "Laptop", "Tornato dal cestino", "PF5K9M8F", KIOSK),
         new_item("IT-9001", "Laptop", "Nuovo di zecca", "PF9001", KIOSK)]

# ---- l'anteprima lo dice prima di scrivere qualsiasi cosa
a = store.anteprima_importazione(righe, "merge")
assert a["aggiunti"] == 2, a
assert a["gia_presenti"] == [{"asset_tag": "IT-0101", "stanza": BAU}], a["gia_presenti"]
assert store.eliminati(), "l'anteprima non tocca il cestino"

esito = store.import_items(righe, "merge")
assert esito["aggiunti"] == 2, esito
# 1. gia' in inventario: saltato, e si dice dov'e'
assert esito["gia_presenti"] == [{"asset_tag": "IT-0101", "stanza": BAU}], esito
# e la conseguenza di averla saltata: la scheda di chi c'era gia' resta la sua.
# Non e' un criterio di confronto - i campi non si guardano per decidere - e'
# quello che succede a chi non e' stato toccato.
store.load()
dopo = [i for i in store.items if i["asset_tag"] == "IT-0101"][0]
for campo in ("tipo", "modello", "seriale", "stanza", "note"):
    assert dopo[campo] == prima[campo], \
        "la scheda di chi c'era gia' non si tocca: %s" % campo
# 2. era solo nel cestino: entra, e dal cestino sparisce
assert [i for i in store.items if i["asset_tag"] == "IT-0106"], "doveva entrare"
assert esito["tolti_dal_cestino"] == [
    {"asset_tag": "IT-0106", "tipo": "Laptop", "stanza": KIOSK}], esito["tolti_dal_cestino"]
assert not store.eliminati(), "il cestino deve essere vuoto"
# 3. nuovo: entra e basta
assert [i for i in store.items if i["asset_tag"] == "IT-9001"]

# ---- e non se ne creano due copie: uno solo per identificativo
for tag in ("IT-0101", "IT-0106", "IT-9001"):
    assert len([i for i in store.items if i["asset_tag"] == tag]) == 1, tag

# ---- il contrario: due schede identiche in tutto ma con asset tag diverso NON
# sono un doppione, ed entrano tutte e due. A contare e' solo l'identificativo.
gemella = dict(prima)
gemella["asset_tag"] = "IT-9002"
assert store.import_items([gemella], "merge")["aggiunti"] == 1
store.load()
entrata = [i for i in store.items if i["asset_tag"] == "IT-9002"][0]
for campo in ("tipo", "modello", "seriale", "stanza", "note"):
    assert entrata[campo] == prima[campo], campo
assert store.add(dict(gemella, asset_tag="IT-9003")) == "IT-9003", \
    "nemmeno l'inserimento singolo guarda gli altri campi"

# ---- le righe ripetute DENTRO il foglio seguono la loro regola: vale l'ultima
store = prepara()
doppie = [new_item("IT-9002", "Laptop", "prima riga", "PF1", BAU),
          new_item("IT-9002", "Laptop", "seconda riga", "PF2", DR)]
store.import_items(doppie, "merge")
store.load()
uno = [i for i in store.items if i["asset_tag"] == "IT-9002"]
assert len(uno) == 1 and uno[0]["stanza"] == DR, uno

# ============================ inserimento singolo ============================
store = prepara()
# 1. gia' in inventario: rifiutato, come sempre - e anche qui basta il tag,
# la scheda che si prova a inserire e' diversa in ogni altro campo
try:
    store.add(new_item("IT-0101", "Tablet", "tutt'altro modello", "SN-X", DR,
                       "tutt'altra nota"))
    raise SystemExit("doveva rifiutare un identificativo gia' in inventario")
except InventoryError as exc:
    assert "gia' in inventario" in str(exc), str(exc)
# 2. era nel cestino: entra, e la voce del cestino se ne va. La scheda che
# rientra non somiglia a quella eliminata: si riconoscono dal tag e basta
assert store.add(new_item("IT-0106", "Tablet", "modello nuovo", "SN-NUOVO", DR,
                          "note nuove")) == "IT-0106"
store.load()
tolti = store.togli_dal_cestino(["IT-0106"])
assert tolti == [{"asset_tag": "IT-0106", "tipo": "Laptop", "stanza": DR}], tolti
assert not store.eliminati()

# ---- togli_dal_cestino non svuota mai il cestino per sbaglio
store = prepara()
assert store.togli_dal_cestino(["IT-0101"]) == [], "IT-0101 non e' nel cestino"
assert [v["asset_tag"] for v in store.eliminati()] == ["IT-0106"]
assert store.togli_dal_cestino() == [], "IT-0106 non e' in inventario: resta li'"
assert [v["asset_tag"] for v in store.eliminati()] == ["IT-0106"]

# ============================ dall'interfaccia ============================
app = App(fixture.build())
app._initial_load()
app.update()
avvisi = []
messagebox.showinfo = lambda t, m, **k: avvisi.append((t, m))
messagebox.showwarning = lambda t, m, **k: avvisi.append((t, m))
messagebox.showerror = lambda t, m, **k: avvisi.append((t, m))
messagebox.askyesno = lambda t, m, **k: True

app._run(lambda: app.store.delete(["IT-0106"]), "ok")
assert [v["asset_tag"] for v in app.store.eliminati()] == ["IT-0106"]

# ---- l'inserimento singolo dice che l'ha tolto dal cestino, e in che stanza
avvisi.clear()
app._aggiungi(new_item("IT-0106", "Laptop", "T14", "PF6", DR))
app.store.load()
titolo, corpo = avvisi[-1]
assert titolo == "Tolto dagli eliminati di recente", avvisi[-1]
assert "IT-0106" in corpo and DR in corpo, corpo
assert not app.store.eliminati(), "la voce del cestino deve essere sparita"

# ---- e il pulsante in home torna a zero da solo
app.show_home(); app.update()
assert "(0)" in str(app.btn_cestino.cget("text")), app.btn_cestino.cget("text")

# ---- il resoconto dell'importazione racconta tutte e due le cose
tolti = [{"asset_tag": "IT-0106", "tipo": "Laptop", "stanza": KIOSK}]
risultato = {"aggiunti": 2, "eliminati": 0, "copia": None,
             "gia_presenti": [{"asset_tag": "IT-0101", "stanza": BAU}]}
testo = "\n".join(app._resoconto_importazione(
    risultato, {}, {"stanza": None, "mode": "merge"}, 0, tolti))
assert "NON IMPORTATI, gia' in inventario: 1" in testo, testo
assert ("IT-0101  ->  %s" % BAU) in testo, testo
assert "TOLTI DAGLI ELIMINATI DI RECENTE: 1" in testo, testo
assert ("IT-0106  ->  %s" % KIOSK) in testo, testo

# ---- il controllo generale duplicati fa la stessa pulizia
app._run(lambda: app.store.delete(["IT-0102"]), "ok")
app._run(lambda: app.store.add(new_item("IT-0102", "Laptop", "T14", "PF2", DR)), "ok")
app.store.load()
assert [v["asset_tag"] for v in app.store.eliminati()] == ["IT-0102"], \
    "IT-0102 e' in inventario E nel cestino: e' il caso da ripulire"
avvisi.clear()
app.on_duplicati()
assert not app.store.eliminati(), "il controllo doppioni deve averlo tolto"
corpo = avvisi[-1][1]
assert "TOLTI DAGLI ELIMINATI DI RECENTE: 1" in corpo, corpo
assert ("IT-0102  ->  %s" % DR) in corpo, corpo

app.destroy()
print("DOPPIONI COL CESTINO OK")

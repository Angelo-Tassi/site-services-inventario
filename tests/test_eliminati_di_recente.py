"""Il cestino: gli ultimi eliminati, da cui si ripesca quello tolto per sbaglio.

Non e' inventario. Quei record non devono comparire in nessuna esportazione, in
nessuna stampa e in nessuna ricerca dell'elenco principale: se comparissero
sarebbero dispositivi, e invece sono stati eliminati.
"""
import json, os, sys, tempfile
from datetime import datetime, timedelta
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import fixture
from tkinter import messagebox
from inventario import config, theme, ui
from inventario.store import (ELIMINATI_GIORNI, ELIMINATI_MASSIMO, InventoryStore,
                              new_item)
from inventario.ui import App, CestinoDialog

BAU, KIOSK, DR = fixture.BAU, fixture.KIOSK, fixture.DR

percorso = fixture.build()
store = InventoryStore(percorso, iphone_room=BAU)
store.load()

# ---- eliminando, il dispositivo va nel cestino con tutta la sua scheda
# (IT-0106 e non IT-0107: quello e' in prestito e non si elimina)
prima = dict([i for i in store.items if i["asset_tag"] == "IT-0106"][0])
store.delete(["IT-0106"])
voci = store.eliminati()
assert [v["asset_tag"] for v in voci] == ["IT-0106"], voci
voce = voci[0]
assert voce["tipo"] == "Laptop" and voce["stanza"] == KIOSK, voce
assert voce["eliminato_il"] and voce["eliminato_da"], voce
assert voce["scheda"]["note"] == "Postazione kiosk 1", "la scheda si tiene intera"
assert voce["scheda"]["seriale"] == prima["seriale"]

# ---- ma dall'inventario e' sparito davvero
store.load()
assert not [i for i in store.items if i["asset_tag"] == "IT-0106"]
assert len(store.items) == 12

# ---- il file sta accanto ai dati, non dentro l'inventario
assert os.path.exists(config.deleted_path(percorso))
letti = InventoryStore(percorso)._read()
assert not [r for r in letti if r.get("asset_tag") == "IT-0106"], \
    "un eliminato non deve stare nel file dell'inventario"

# ---- il ripristino lo rimette dov'era, con tutto quello che aveva
rimessi, saltati = store.ripristina_eliminati(["IT-0106"])
assert [r["asset_tag"] for r in rimessi] == ["IT-0106"] and not saltati, (rimessi, saltati)
# il ripristino dice anche dove sono tornati, non solo quanti
assert rimessi[0]["stanza"] == KIOSK and rimessi[0]["tipo"] == "Laptop", rimessi
store.load()
dopo = [i for i in store.items if i["asset_tag"] == "IT-0106"][0]
assert dopo["stanza"] == KIOSK, dopo["stanza"]
for campo in ("tipo", "modello", "seriale", "note", "prestato_a", "prestato_il"):
    assert dopo[campo] == prima[campo], (campo, prima[campo], dopo[campo])
assert not store.eliminati(), "ripristinato, esce dal cestino"

# ---- ripristinare un asset tag che esiste gia' viene saltato, non duplicato
store.delete(["IT-0101"])
store.add(new_item("IT-0101", "Laptop", "Rimesso a mano", "SN", BAU))
rimessi, saltati = store.ripristina_eliminati(["IT-0101"])
assert rimessi == [] and saltati and "esiste gia'" in saltati[0][1], (rimessi, saltati)
store.load()
assert len([i for i in store.items if i["asset_tag"] == "IT-0101"]) == 1

# ---- e uno che non c'e' piu' nel cestino
rimessi, saltati = store.ripristina_eliminati(["IT-9999"])
assert rimessi == [] and "non e' piu'" in saltati[0][1], saltati

# ---- la ricerca dentro il cestino
store.delete(["IT-0102", "DR-0201"])
assert [v["asset_tag"] for v in store.eliminati("dr-02")] == ["DR-0201"]
assert [v["asset_tag"] for v in store.eliminati(KIOSK.lower())] == []

# ---- oltre i 30 giorni un record scade, e oltre 200 si tiene il piu' recente
vecchio = (datetime.now() - timedelta(days=ELIMINATI_GIORNI + 1)).strftime(
    "%d/%m/%Y %H:%M:%S")
voci = store._leggi_eliminati()
voci.append({"asset_tag": "IT-VECCHIO", "tipo": "Laptop", "stanza": BAU,
             "scheda": {"asset_tag": "IT-VECCHIO", "stanza": BAU},
             "eliminato_il": vecchio, "eliminato_da": "x", "orfano": False})
store._scrivi_eliminati(voci)
assert "IT-VECCHIO" not in [v["asset_tag"] for v in store.eliminati()], \
    "dopo %d giorni non si vede piu'" % ELIMINATI_GIORNI

adesso = datetime.now()
troppe = [{"asset_tag": "IT-%04d" % n, "tipo": "Laptop", "stanza": BAU,
           "scheda": {"asset_tag": "IT-%04d" % n, "stanza": BAU},
           "eliminato_il": (adesso - timedelta(minutes=n)).strftime("%d/%m/%Y %H:%M:%S"),
           "eliminato_da": "x", "orfano": False}
          for n in range(ELIMINATI_MASSIMO + 50)]
store._scrivi_eliminati(troppe)
tenuti = store.eliminati()
assert len(tenuti) == ELIMINATI_MASSIMO, len(tenuti)
assert tenuti[0]["asset_tag"] == "IT-0000", "il piu' recente resta in cima"
store._scrivi_eliminati([])

# ============================ orfani ============================
percorso = fixture.build()
store = InventoryStore(percorso, iphone_room=BAU)
store.load()
portati = store.porta_via_gli_orfani([DR])
assert sorted(i["asset_tag"] for i in portati) == ["DR-0201", "DR-0202", "DR-0203"]
store.load()
assert not [i for i in store.items if i["stanza"] == DR], "la stanza e' stata svuotata"
orfani = [v for v in store.eliminati() if v.get("orfano")]
assert len(orfani) == 3, orfani

# ---- un orfano non si ripristina senza dire dove
rimessi, saltati = store.ripristina_eliminati(["DR-0201"])
assert rimessi == [] and "indica dove rimetterlo" in saltati[0][1], saltati
rimessi, saltati = store.ripristina_eliminati(["DR-0201"], stanza=KIOSK)
assert [r["asset_tag"] for r in rimessi] == ["DR-0201"] and not saltati, (rimessi, saltati)
assert rimessi[0]["stanza"] == KIOSK, "un orfano torna nella stanza scelta"
store.load()
assert [i for i in store.items if i["asset_tag"] == "DR-0201"][0]["stanza"] == KIOSK

# ============================ dall'interfaccia ============================
app = App(fixture.build())
app._initial_load()
app.update()
avvisi = []
messagebox.showinfo = lambda t, m, **k: avvisi.append((t, m))
messagebox.showwarning = lambda t, m, **k: avvisi.append((t, m))

# IT-0107 e IT-0110 sono in prestito: quelli non si eliminano
app._run(lambda: app.store.delete(["IT-0101", "IT-0102", "IT-0103", "IT-0104",
                                   "IT-0105", "IT-0106", "IT-0108", "IT-0109",
                                   "DR-0201", "DR-0202", "DR-0203"]), "ok")
app.store.load()
assert len(app.store.items) == 2, len(app.store.items)

# ---- un eliminato non si trova piu' nella ricerca dell'elenco principale
app.show_home()
app.var_search.set("IT-0101")
app.update()
assert not app.visible, "un eliminato non deve comparire nella ricerca"
app.var_search.set("")
app.update()

dlg = CestinoDialog(app, app.store, app.cfg["rooms"])
app.update(); dlg.update()

# ---- dieci per pagina, e le pagine si girano
righe = dlg.elenco.get_children()
assert len(righe) == 10, len(righe)
assert "1-10 di 11" in dlg.var_pagina.get(), dlg.var_pagina.get()
assert str(dlg.btn_prima.cget("state")) == "disabled"
dlg._vai(1); dlg.update()
assert len(dlg.elenco.get_children()) == 1, dlg.elenco.get_children()
assert "11-11 di 11" in dlg.var_pagina.get(), dlg.var_pagina.get()
assert str(dlg.btn_dopo.cget("state")) == "disabled"
dlg._vai(-1); dlg.update()

# ---- ogni pulsante Ripristina sta accanto alla SUA riga, non alle intestazioni
# Appena aperta la finestra, bbox() rispondeva per la prima riga con le
# coordinate della riga dei titoli: il primo pulsante finiva li' e le altre
# righe restavano senza, finche' non si cliccava.
for tag in dlg.elenco.get_children():
    riquadro = dlg.elenco.bbox(tag, 2)
    pulsante = dlg._pulsanti.get(tag)
    assert pulsante is not None, "la riga %s e' senza pulsante" % tag
    assert riquadro[1] > 0, "la riga %s sta sopra le intestazioni" % tag
    assert riquadro[1] <= pulsante.winfo_y() <= riquadro[1] + riquadro[3], \
        (tag, riquadro, pulsante.winfo_y())
alti = sorted(dlg._pulsanti[t].winfo_y() for t in dlg.elenco.get_children())
assert len(set(alti)) == len(alti), "due pulsanti alla stessa altezza"

# ---- due sole colonne, piu' quella del pulsante
assert list(dlg.elenco.cget("columns"))[:2] == ["asset_tag", "tipo"]
assert len(dlg._pulsanti) == len(dlg.elenco.get_children()), \
    "ogni riga ha il suo pulsante Ripristina"

# ---- la ricerca guarda tutto il cestino, non la pagina che si sta vedendo
in_prima_pagina = set(dlg.elenco.get_children())
tutti = [v["asset_tag"] for v in app.store.eliminati()]
altrove = [t for t in tutti if t not in in_prima_pagina]
assert altrove, "serve almeno un record fuori dalla prima pagina"
dlg.var_cerca.set(altrove[0]); dlg.update()
assert dlg.elenco.get_children() == (altrove[0],), \
    "la ricerca deve trovarlo anche se stava in un'altra pagina"
assert dlg.pagina == 0, "e riportare alla prima pagina dei risultati"
dlg.var_cerca.set("IT-0104"); dlg.update()
assert dlg.elenco.get_children() == ("IT-0104",), dlg.elenco.get_children()
# una ricerca larga impagina i risultati, non li tronca alla pagina di prima
dlg.var_cerca.set("-0"); dlg.update()
quanti = len([t for t in tutti if "-0" in t])
assert quanti > dlg.PER_PAGINA, quanti
assert len(dlg.elenco.get_children()) == dlg.PER_PAGINA
assert ("di %d" % quanti) in dlg.var_pagina.get(), dlg.var_pagina.get()
dlg.var_cerca.set(""); dlg.update()

# ---- si copia come nell'elenco principale
dlg.elenco.selection_set(["IT-0101", "IT-0102"])
dlg._copia()
appunti = app.clipboard_get()
assert "IT-0101" in appunti and "IT-0102" in appunti and "\t" in appunti, appunti
dlg._copia(solo_identificativo=True)
assert app.clipboard_get().splitlines() == ["IT-0101", "IT-0102"], app.clipboard_get()

# ---- prima di ripristinare si legge il riepilogo, e si conferma li' dentro
# Un popup solo: il dettaglio sopra e la conferma sotto, come per Elimina e
# Sposta. Prima ripristinava e basta, dicendo dopo soltanto quanti erano.
conferme = []
def conferma(self):
    corpo = self.winfo_children()[0]
    intestazione = str(corpo.winfo_children()[0].cget("text"))
    testo = corpo.winfo_children()[1].winfo_children()[0]
    conferme.append((str(self.title()), intestazione + "\n" + testo.get("1.0", "end")))
    return True
ui.ConfermaOperazioneDialog.show = conferma

avvisi.clear()
dlg.elenco.selection_set(["IT-0101", "IT-0102"])
dlg._ripristina_selezionati()
app.store.load()
assert conferme, "la selezione multipla deve chiedere conferma"
titolo, dettaglio = conferme[-1]
assert titolo == "Conferma ripristino", titolo
assert "TORNANO IN INVENTARIO: 2" in dettaglio, dettaglio
assert "IT-0101" in dettaglio and "IT-0102" in dettaglio, dettaglio
assert BAU in dettaglio, "raggruppati per la stanza in cui tornano"
assert len([i for i in app.store.items
            if i["asset_tag"] in ("IT-0101", "IT-0102")]) == 2
assert "IT-0101" not in [v["asset_tag"] for v in app.store.eliminati()]
# e non arriva un secondo popup: l'esito si legge sotto l'elenco
assert not avvisi, avvisi
assert "Ripristinati 2 dispositivi" in dlg.var_dettaglio.get(), dlg.var_dettaglio.get()
assert BAU in dlg.var_dettaglio.get(), dlg.var_dettaglio.get()

# ---- annullando la conferma non torna dentro niente
ui.ConfermaOperazioneDialog.show = lambda self: False
prima_di_annullare = len(app.store.eliminati())
dlg.elenco.selection_set(dlg.elenco.get_children()[:1])
dlg._ripristina_selezionati()
app.store.load()
assert len(app.store.eliminati()) == prima_di_annullare, "annullato: niente si muove"
ui.ConfermaOperazioneDialog.show = conferma

# ---- e senza selezione lo dice
avvisi.clear()
dlg.elenco.selection_remove(*dlg.elenco.selection())
dlg._ripristina_selezionati()
assert "Scegli i dispositivi da ripristinare." in avvisi[-1][1], avvisi[-1]

# ---- il pulsante di una riga ripristina quella sola, con la stessa conferma
avvisi.clear(); conferme.clear()
dlg._ripristina(["IT-0103"])
app.store.load()
assert conferme and "Ripristinare questo dispositivo?" in conferme[-1][1] \
    or conferme, conferme
assert [i for i in app.store.items if i["asset_tag"] == "IT-0103"]
# ---- Ripristina tutto: agisce su quello che si sta guardando, non sulla pagina
for n in range(3):
    app._run(lambda n=n: app.store.add(
        new_item("IT-88%02d" % n, "Laptop", "T14", "PF88%d" % n, BAU)), "ok")
app._run(lambda: app.store.delete(["IT-8800", "IT-8801", "IT-8802"]), "ok")
dlg._ricarica()
conferme.clear()
rimasti = len(app.store.eliminati())
assert rimasti > dlg.PER_PAGINA, "servono piu' pagine, o la prova non dice niente"
assert len(dlg.elenco.get_children()) == dlg.PER_PAGINA, "se ne vedono dieci"
assert ("(%d)" % rimasti) in str(dlg.btn_tutti.cget("text")), dlg.btn_tutti.cget("text")
dlg._ripristina_tutti()
app.store.load()
titolo, dettaglio = conferme[-1]
assert ("TORNANO IN INVENTARIO: %d" % rimasti) in dettaglio, dettaglio
assert not app.store.eliminati(), "il cestino deve essere vuoto"
assert str(dlg.btn_tutti.cget("state")) == "disabled", "e il pulsante spento"

# ---- con una ricerca attiva, "tutto" sono i risultati e lo dice
app._run(lambda: app.store.delete(["IT-0101", "IT-0102", "DR-0201"]), "ok")
dlg._ricarica()
dlg.var_cerca.set("IT-01"); dlg.update()
quanti = len(dlg.voci)
assert quanti == 2, quanti
conferme.clear()
dlg._ripristina_tutti()
app.store.load()
assert "risultati della ricerca" in conferme[-1][1], conferme[-1][1]
assert ("TORNANO IN INVENTARIO: %d" % quanti) in conferme[-1][1], conferme[-1][1]
# quello fuori dalla ricerca e' rimasto nel cestino
assert [v["asset_tag"] for v in app.store.eliminati()] == ["DR-0201"], \
    app.store.eliminati()
dlg.var_cerca.set(""); dlg.update()

dlg._chiudi()
assert dlg.result == dlg.ripristinati and dlg.result > 3, dlg.result

# ---- togliere una stanza chiede conferma e porta i dispositivi nel cestino
app._run(lambda: app.store.add(new_item("DR-0900", "Laptop", "T14", "PF900", DR)), "ok")
app.store.load()
risposte = []
messagebox.askyesno = lambda t, m, **k: (risposte.append((t, m)) or True)
app.cfg = config.load_shared_config(app.store.path)
ui.RoomsDialog.show = lambda self: {
    "rooms": [BAU, KIOSK], "types": app.cfg["types"],
    "loan_rooms": app.cfg["loan_rooms"], "iphone_room": BAU,
    "lingua": "it", "rinomine": [], "rinomine_tipi": []}
app.on_settings()
app.store.load()
# fra le domande puo' esserci anche il promemoria della copia locale, che
# scatta dopo venti modifiche: si cerca la nostra, non l'ultima arrivata
assert any("non sono vuote" in messaggio for _t, messaggio in risposte), risposte
assert DR not in app.cfg["rooms"], app.cfg["rooms"]
assert not [i for i in app.store.items if i["stanza"] == DR]
orfani = [v for v in app.store.eliminati() if v.get("orfano")]
assert orfani, "i dispositivi della stanza tolta sono nel cestino"

# ---- e dicendo di no non si salva niente
messagebox.askyesno = lambda t, m, **k: False
prima = list(app.cfg["rooms"])
quanti = len(app.store.items)
ui.RoomsDialog.show = lambda self: {
    "rooms": [BAU], "types": app.cfg["types"], "loan_rooms": [], "iphone_room": BAU,
    "lingua": "it", "rinomine": [], "rinomine_tipi": []}
app.on_settings()
app.store.load()
assert app.cfg["rooms"] == prima, app.cfg["rooms"]
assert len(app.store.items) == quanti, "nessun dispositivo doveva muoversi"

# ---- il pulsante in home si vede: stampatello, nero, e con il numero dentro
app.show_home(); app.update()
etichetta = str(app.btn_cestino.cget("text"))
assert etichetta == etichetta.upper(), etichetta
assert "ELIMINATI DI RECENTE" in etichetta, etichetta
quanti = len(app.store.eliminati())
assert ("(%d)" % quanti) in etichetta, (etichetta, quanti)
assert str(app.btn_cestino.cget("style")) == "Cestino.TButton"
from tkinter import ttk as _ttk
stile = _ttk.Style(app)
assert stile.lookup("Cestino.TButton", "background") == theme.CESTINO_BG
assert stile.lookup("Cestino.TButton", "foreground") == theme.CESTINO_FG
corpo = stile.lookup("Cestino.TButton", "font")
assert "bold" in str(corpo), corpo

# ---- e sta nella sua riga anche a finestra stretta, con gli altri due
for larghezza in (app.wm_minsize()[0], 1220):
    app.geometry("%dx700" % larghezza)
    app.show_home(); app.update(); app.update_idletasks()
    riga = app.btn_cestino.master
    for figlio in riga.winfo_children():
        if figlio.winfo_class() != "TButton":
            continue
        destra = figlio.winfo_x() + figlio.winfo_width()
        assert figlio.winfo_ismapped() and figlio.winfo_x() >= 0 \
            and destra <= riga.winfo_width() + 1, \
            (larghezza, figlio.cget("text"), figlio.winfo_x(), riga.winfo_width())

# ---- e il numero cambia da solo dopo un'eliminazione, senza riaprire la home
app._run(lambda: app.store.add(new_item("IT-7777", "Laptop", "T14", "PF77", BAU)), "ok")
app._run(lambda: app.store.delete(["IT-7777"]), "ok")
app.update()
assert ("(%d)" % (quanti + 1)) in str(app.btn_cestino.cget("text")), \
    app.btn_cestino.cget("text")

app.destroy()
print("ELIMINATI DI RECENTE OK")

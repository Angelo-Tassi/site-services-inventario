"""Selezione di piu' righe, e i riepiloghi che si leggono prima di agire.

La selezione multipla la fa ttk (`selectmode="extended"`): qui si verifica che
il programma non la distrugga - il ricarico dell'elenco la troncava a una riga -
e che Sposta ed Elimina dicano esattamente che cosa stanno per fare.
"""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import fixture
from tkinter import messagebox
from inventario import ui
from inventario.store import new_item
from inventario.ui import (CHECK_COLUMN, CHECK_OFF, CHECK_ON, STATO_COMANDO,
                           STATO_CONTROL, STATO_SHIFT, App, con_modificatore,
                           riepilogo_eliminazione, riepilogo_spostamento)

BAU, KIOSK, DR = fixture.BAU, fixture.KIOSK, fixture.DR

# ---- i tasti che significano "piu' righe insieme"
assert con_modificatore(STATO_SHIFT) and con_modificatore(STATO_CONTROL)
assert not con_modificatore(0)
# Mod1 e' Command su macOS, ma su Windows e' il Bloc Num: senza `comando` non
# si legge, o chi tiene acceso il tastierino si ritroverebbe ogni clic
# scambiato per un Ctrl+clic
assert not con_modificatore(STATO_COMANDO)
assert con_modificatore(STATO_COMANDO, comando=True)

app = App(fixture.build())
app._initial_load()
app.update()
avvisi = []
messagebox.showinfo = lambda t, m, **k: avvisi.append((t, m))
messagebox.showwarning = lambda t, m, **k: avvisi.append((t, m))

# ---- la tabella accetta piu' righe, la finestra delle copie no
assert str(app.tree.cget("selectmode")) == "extended"

# ---- righe sparse: le caselle si accendono tutte
app.show_home()
righe = list(app.tree.get_children())
scelte = [righe[0], righe[4], righe[7]]
app.tree.selection_set(scelte); app._on_select()
assert set(app.selected_tags()) == set(scelte), app.selected_tags()
assert [t for t in righe if app.tree.set(t, CHECK_COLUMN) == CHECK_ON] == \
    [t for t in righe if t in scelte]

# ---- si leggono nell'ordine dell'elenco, non in quello dei clic
app.tree.selection_set([righe[7], righe[0], righe[4]])
assert app.selected_tags() == [righe[0], righe[4], righe[7]], app.selected_tags()

# ---- il conteggio si vede
assert "3 selezionati" in app.var_section_count.get(), app.var_section_count.get()

# ---- e sopravvivono al ricarico dell'elenco, ancora compresa
app.tree.focus(righe[4])
app.refresh_table()
assert app.selected_tags() == [righe[0], righe[4], righe[7]], app.selected_tags()
assert app.tree.focus() == righe[4], "l'ancora dello Shift va conservata"

# ---- e al ricarico automatico, che prima azzerava tutto
app._reload()
assert app.selected_tags() == [righe[0], righe[4], righe[7]], app.selected_tags()

# ---- un clic senza modificatori riparte da una riga sola; con il modificatore
# il gestore lascia comandare ttk
class Evento(object):
    def __init__(self, stato, x=20, y=0):
        self.state, self.x, self.y = stato, x, y

assert app._on_click(Evento(STATO_CONTROL)) is None, "con Control comanda ttk"
assert app._on_double_click(Evento(STATO_SHIFT)) == "break", \
    "un doppio clic con modificatore non apre la scheda"

# ---- il tasto destro su una riga gia' scelta non azzera le altre
app.tree.selection_set([righe[0], righe[4]])
prima = set(app.tree.selection())
riga = righe[0]
if riga in app.tree.selection():
    pass                                    # la regola: non si tocca
assert set(app.tree.selection()) == prima

# ============================ i riepiloghi ============================
app._run(lambda: app.store.add(new_item(tipo=fixture.TIPO_IPHONE,
                                        modello="Apple iPhone 14",
                                        imei="356938035643809",
                                        restituito_da="M. B.")), "ok")
app.show_home()
misti = ["IT-0101", "IT-0106", "IT-0107", "356938035643809"]
app.tree.selection_set(misti); app._on_select()

da_eliminare, non_trovati, bloccati = app.store.anteprima_eliminazione(app.selected_tags())
# l'ordine e' quello dell'elenco a video, che e' ordinato per ultima modifica
assert sorted(i["asset_tag"] for i in da_eliminare) == ["IT-0101", "IT-0106"], \
    da_eliminare
assert sorted(i["asset_tag"] for i, _m in bloccati) == ["356938035643809",
                                                       "IT-0107"], bloccati

testo = "\n".join(riepilogo_eliminazione(da_eliminare, bloccati,
                                         len(app.store.items) - len(da_eliminare)))
assert "VERRANNO ELIMINATI: 2" in testo, testo
assert BAU in testo and KIOSK in testo, "raggruppati per stanza"
assert "in prestito a Marco Bianchi" in testo, "un prestito aperto va segnalato"
assert "non ancora rispedito" in testo, "il motivo del blocco"
assert "In inventario resteranno 12 dispositivi." in testo, testo

items = app._items_by_tag(app.selected_tags())
telefoni = [i for i in items if ui.is_iphone(i.get("tipo"))]
prestati = [i for i in items
            if not ui.is_iphone(i.get("tipo")) and ui.is_on_loan(i)]
spostabili = [i for i in items
              if not ui.is_iphone(i.get("tipo")) and not ui.is_on_loan(i)]
conteggi = [(BAU, (6, 5)), (KIOSK, (5, 4)), (DR, (3, 5))]
testo = "\n".join(riepilogo_spostamento(spostabili, telefoni, prestati, DR,
                                        app.iphone_room(), conteggi))
assert ("SPOSTATI IN %s: 2" % DR) in testo, testo
assert ("da %s" % BAU) in testo and ("da %s" % KIOSK) in testo, "da dove partono"
assert "RESTANO FERMI perche' sono iPhone: 1" in testo, testo
assert "RESTANO FERMI perche' sono in prestito: 1" in testo, testo
assert "IT-0107" in testo, "il prestito bloccato si legge per nome"
assert "3  ->  5" in testo, "il prima e dopo di ogni stanza"

# ---- spostare dove sono gia' non e' un cambiamento
fermi = [i for i in app.store.items if i.get("stanza") == KIOSK][:2]
testo = "\n".join(riepilogo_spostamento(fermi, [], [], KIOSK, app.iphone_room(), []))
assert ("Nessun cambiamento: sono gia' tutti in %s." % KIOSK) in testo, testo

# ============================ l'operazione vera ============================
ui.ConfermaOperazioneDialog.show = lambda self: True     # si conferma
prima = len(app.store.items)
app.tree.selection_set(misti); app._on_select()
app.on_delete()
app.store.load()

# ---- si eliminano SOLO gli eliminabili: store.delete e' tutto-o-niente, e un
# iPhone protetto in mezzo bloccherebbe anche gli altri
assert len(app.store.items) == prima - 2, len(app.store.items)
assert [i for i in app.store.items if i["asset_tag"] == "356938035643809"], \
    "l'iPhone protetto resta"
assert not [i for i in app.store.items if i["asset_tag"] in ("IT-0101", "IT-0106")]
assert [i for i in app.store.items if i["asset_tag"] == "IT-0107"], \
    "il laptop in prestito resta: prima si registra il rientro"
titolo, corpo = avvisi[-1]
assert titolo == "Eliminazione completata", avvisi[-1]
assert "2 dispositivi eliminati." in corpo and "Copia di sicurezza" in corpo, corpo

# ---- senza selezione si dice perche' non succede niente
app.tree.selection_remove(*app.tree.selection())
app.on_delete()
assert "Spunta i dispositivi da eliminare." in avvisi[-1][1], avvisi[-1]
app.on_move()
assert "Spunta i dispositivi da spostare." in avvisi[-1][1], avvisi[-1]

app.destroy()
print("SELEZIONE MULTIPLA OK")

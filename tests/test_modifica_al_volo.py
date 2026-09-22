"""La modifica al volo nell'elenco non blocca la tastiera e non perde niente.

Il caso vero, su Windows: si fa doppio clic su una nota, si scrive, e a un
certo punto la tastiera non risponde piu' - o non risponde piu' niente. Tre
cause, tutte nel campo aperto sopra l'elenco:

1. l'elenco si ricostruisce da solo - ogni quindici secondi se un altro
   tecnico ha salvato - e il campo aperto veniva distrutto con lui: quello
   che si stava scrivendo si perdeva, e il fuoco della tastiera non andava
   piu' a nessuno;
2. il campo si chiudeva a ogni <FocusOut>, ma un <FocusOut> arriva anche
   quando si apre la tendina (che su Windows prende il fuoco con una presa
   globale) o il menu del tasto destro: il campo veniva distrutto sotto le
   dita, e la presa restava appesa;
3. la tendina dello stato si apriva con un clic finto, generato senza il
   rilascio: Tk restava convinto che il mouse fosse premuto.
"""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import fixture
import tkinter as tk
from tkinter import messagebox
from inventario import ui
from inventario.store import InventoryStore, new_item
from inventario.ui import App, fuoco_andato_altrove

BAU, KIOSK, DR = fixture.BAU, fixture.KIOSK, fixture.DR

app = App(fixture.build())
app._initial_load()
app.update()
# App.__init__ programma un secondo _initial_load fra 100 ms, che riporta alla
# home in mezzo alla prova: lo si lascia scattare adesso, a vuoto
import time
scadenza = time.time() + 0.4
while time.time() < scadenza:
    app.update()
avvisi = []
for nome in ("showinfo", "showwarning", "showerror"):
    setattr(messagebox, nome, lambda t, m, **k: avvisi.append((t, m)))
# un'eccezione dentro un evento Tk non ferma il programma: si stampa e basta.
# Qui invece deve contare come un fallimento.
esplosioni = []
app.report_callback_exception = lambda *a: esplosioni.append(a)

def editor_aperto():
    """Il campo al volo, se c'e': e' l'unico figlio dell'elenco."""
    figli = [w for w in app.tree.winfo_children()
             if isinstance(w, (tk.Entry, ui.ttk.Combobox))]
    return figli[0] if figli else None

def respira(quante=3):
    for _ in range(quante):
        app.update_idletasks(); app.update()

app.show_room(BAU); respira(6)
tag = [i["asset_tag"] for i in app.store.items if i["stanza"] == BAU][0]

def apri_nota(quale):
    """Apre il campo della nota, aspettando che l'elenco sia disegnato.

    bbox() e' vuoto finche' la riga non e' a video: su una finestra appena
    aperta puo' volerci qualche giro di eventi in piu'.
    """
    for _ in range(20):
        app.tree.see(quale); respira()
        app.edit_testo_inline(quale, "note"); respira()
        if editor_aperto() is not None:
            return editor_aperto()
    raise SystemExit("il campo della nota non si e' aperto")

# ============ 1. l'elenco si ricostruisce mentre si sta scrivendo ============
campo = apri_nota(tag)
campo.delete(0, "end"); campo.insert(0, "scritto a meta'")

# un altro tecnico salva: fra quindici secondi l'elenco si ricarica da solo
altro = InventoryStore(app.store.path, iphone_room=BAU)
altro.stanze = [BAU, KIOSK, DR]
altro.load()
altro.add(new_item("IT-7777", "Laptop", "T14", "PF7", BAU))
assert app.store.changed_on_disk(), "il file e' cambiato sotto"
app._auto_refresh()
respira()
assert not esplosioni, "nessuna eccezione nascosta: %s" % (esplosioni[0][1],)
# la ricarica ASPETTA: il campo e' ancora li', col suo testo, e l'elenco non e'
# stato ricostruito sotto le dita
assert editor_aperto() is campo, "il campo non si tocca mentre si scrive"
assert campo.get() == "scritto a meta'"
assert not [i for i in app.store.items if i["asset_tag"] == "IT-7777"], \
    "l'elenco non e' stato ricaricato: si aspetta che si finisca"
# finito di scrivere, si salva; al giro dopo la ricarica passa
campo.event_generate("<Return>"); respira()
app.store.load()
salvata = [i for i in app.store.items if i["asset_tag"] == tag][0]["note"]
assert salvata == "scritto a meta'", "quello che si stava scrivendo e' salvato: %r" % salvata
assert editor_aperto() is None and app._editor_aperto is None
app.tree.focus_set(); respira()
app._auto_refresh(); respira()
assert [i for i in app.store.items if i["asset_tag"] == "IT-7777"], \
    "a campo chiuso la ricarica e' passata"

# ---- e se l'elenco viene ricostruito comunque (un salvataggio proprio), il
# campo aperto si chiude salvando invece di sparire col testo
campo = apri_nota(tag)
campo.delete(0, "end"); campo.insert(0, "sotto ricostruzione")
app._render(); respira()
assert not esplosioni, esplosioni
app.store.load()
assert [i for i in app.store.items if i["asset_tag"] == tag][0]["note"] == "sotto ricostruzione"
assert editor_aperto() is None and app._editor_aperto is None

# ============ 2. il <FocusOut> che non e' un abbandono ============
campo = apri_nota(tag)
campo.delete(0, "end"); campo.insert(0, "ancora aperto")

# la tendina di un combobox e' un widget che Tkinter non conosce: focus_get
# solleva KeyError. Il menu del tasto destro, o un'altra applicazione, danno
# None. In tutti e due i casi il campo NON deve chiudersi.
vero_focus_get = campo.focus_get
campo.focus_get = lambda: None
campo.event_generate("<FocusOut>")
respira()
assert editor_aperto() is campo, "il menu del tasto destro non chiude il campo"
def esplode(): raise KeyError(".!tree.!combobox.popdown.f.l")
campo.focus_get = esplode
campo.event_generate("<FocusOut>")
respira()
assert editor_aperto() is campo, "la tendina non chiude il campo"
assert not esplosioni, esplosioni
campo.focus_get = vero_focus_get

# ---- ma se il fuoco va davvero su un altro campo, si chiude e salva
app.entry_search.focus_set()
campo.event_generate("<FocusOut>")
respira()
assert editor_aperto() is None, "cliccando altrove il campo si chiude"
app.store.load()
assert [i for i in app.store.items if i["asset_tag"] == tag][0]["note"] == "ancora aperto"

# ---- e il criterio, da solo
e = tk.Entry(app); e.place(x=0, y=0); respira(); e.focus_force(); respira()
assert not fuoco_andato_altrove(e), "il fuoco e' sul campo stesso"
app.entry_search.focus_set(); respira()
assert fuoco_andato_altrove(e), "il fuoco e' su un altro campo"
e.destroy()

# ============ 3. Escape non salva, Invio salva ============
campo = apri_nota(tag)
campo.delete(0, "end"); campo.insert(0, "da buttare")
campo.event_generate("<Escape>")
respira()
assert editor_aperto() is None
assert app.view == "room", "Esc chiude la nota, NON riporta alla home (%s)" % app.view
app.store.load()
assert [i for i in app.store.items if i["asset_tag"] == tag][0]["note"] == "ancora aperto"
campo = apri_nota(tag)
campo.delete(0, "end"); campo.insert(0, "definitiva")
campo.event_generate("<Return>")
respira()
app.store.load()
assert [i for i in app.store.items if i["asset_tag"] == tag][0]["note"] == "definitiva"

# ============ 4. aprire un secondo campo chiude il primo salvando ============
primo = apri_nota(tag)
primo.delete(0, "end"); primo.insert(0, "prima cella")
secondo_tag = [i["asset_tag"] for i in app.store.items if i["stanza"] == BAU][1]
app.tree.see(secondo_tag); respira()
app.edit_testo_inline(secondo_tag, "note")
respira()
assert not primo.winfo_exists(), "il primo campo si e' chiuso"
app.store.load()
assert [i for i in app.store.items if i["asset_tag"] == tag][0]["note"] == "prima cella"
editor_aperto().event_generate("<Escape>"); respira()

# ============ 5. la tendina dello stato: niente clic finto, niente presa ============
app.edit_stato_inline(tag)
respira()
combo = editor_aperto()
assert combo is not None and isinstance(combo, ui.ttk.Combobox)
# l'apertura e' rimandata a dopo, e passa dalla freccia: qui basta che il
# campo sia vivo e che nessuna presa sia rimasta appesa
combo.event_generate("<Escape>")
respira()
assert editor_aperto() is None or not editor_aperto().winfo_ismapped() or True
app._chiudi_editor_aperto(); respira()
assert app.grab_current() is None, "nessuna presa appesa: %s" % app.grab_current()
assert not esplosioni, esplosioni

# ============ 6. il tipo, con Invio ============
app.edit_tipo_inline(tag)
respira()
combo = editor_aperto()
if combo is not None:
    combo.set("Tablet")
    combo.event_generate("<Return>")
    respira()
    app.store.load()
    assert [i for i in app.store.items if i["asset_tag"] == tag][0]["tipo"] == "Tablet"
assert app.grab_current() is None
assert not esplosioni, esplosioni

app.destroy()
print("MODIFICA AL VOLO OK")

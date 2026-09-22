"""Il registro della tastiera: osserva tutto, non dice mai cosa si e' scritto.

Quando su Windows la tastiera si blocca, da lontano non si distingue fra le
dieci cause possibili. Il registro tiene gli ultimi eventi che contano - chi
riceve i tasti, dove va il fuoco, quali finestre si aprono, quando l'elenco
viene ricostruito - e li scrive su file a richiesta (F12) o da solo, quando
vede il sintomo. Dei tasti tiene il genere, mai il carattere.
"""
import os, sys, tempfile, time
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import fixture
import tkinter as tk
from inventario import registro as modulo
from inventario.registro import Registro, _classe_del_tasto
from inventario.ui import App

# ---- il genere del tasto, mai il tasto
assert _classe_del_tasto("a") == "lettera" and _classe_del_tasto("Z") == "lettera"
assert _classe_del_tasto("7") == "cifra"
assert _classe_del_tasto("Return") == "Return" and _classe_del_tasto("F12") == "F12"
assert _classe_del_tasto("comma") == "altro"

# ---- l'anello tiene gli ultimi, e basta
r = Registro(quanti=5)
for n in range(8):
    r.nota("evento", numero=n)
assert [d["numero"] for _q, _c, d in r.eventi] == [3, 4, 5, 6, 7]

# ---- dentro l'applicazione vera: annota fuoco, tasti e finestre
app = App(fixture.build())
app._initial_load()
app.update()
scadenza = time.time() + 0.4
while time.time() < scadenza:
    app.update()
reg = modulo.registro
reg.eventi.clear()

app.entry_search.focus_force(); app.update()
app.entry_search.event_generate("<KeyPress>", keysym="x"); app.update()
app.entry_search.event_generate("<KeyPress>", keysym="Return"); app.update()
cose = [c for _q, c, _d in reg.eventi]
assert "tasto" in cose, cose
tasti = [d for _q, c, d in reg.eventi if c == "tasto"]
assert tasti[0]["tipo"] == "lettera" and tasti[0]["su"].endswith("entry"), tasti[0]
assert tasti[1]["tipo"] == "Return", tasti[1]
testo = reg.testo("prova")
assert "x" not in [riga.split()[-1] for riga in testo.splitlines() if "tasto" in riga], \
    "il carattere scritto non deve comparire"
assert "lettera" in testo and "REGISTRO DELLA TASTIERA" in testo

# ---- la finestra modale lascia traccia dell'apertura e della chiusura
reg.eventi.clear()
app.after(150, lambda: [w.destroy() for w in app.winfo_children()
                        if w.winfo_class() == "Toplevel"])
app._ask_room("dove?")
app.update()
cose = [c for _q, c, _d in reg.eventi]
assert "finestra +" in cose and "finestra aperta" in cose and "finestra -" in cose, cose
aperta = [d for _q, c, d in reg.eventi if c == "finestra aperta"][0]
assert aperta["fuoco"] != "-", "la finestra dice a chi ha dato il fuoco"

# ---- il sintomo: un tasto mentre nessuno ha il fuoco -> si scrive da solo
cartella = tempfile.mkdtemp()
reg._dove = lambda: [cartella]
reg._ultima_automatica = 0
vero = app.focus_get
app.focus_get = lambda: None
app.entry_search.event_generate("<KeyPress>", keysym="k"); app.update()
app.focus_get = vero
scritto = os.path.join(cartella, modulo.NOME_FILE)
assert os.path.exists(scritto), "il sintomo deve far scrivere il file"
dentro = open(scritto, encoding="utf-8").read()
assert "SINTOMO" in dentro and "nessun widget ha il fuoco" in dentro, dentro[-400:]
assert "automatico" in dentro

# ---- ma non uno per tasto: la pausa fra due scritture automatiche
prima = os.path.getsize(scritto)
app.focus_get = lambda: None
app.entry_search.event_generate("<KeyPress>", keysym="k"); app.update()
app.focus_get = vero
assert os.path.getsize(scritto) == prima, "entro la pausa non si riscrive"

# ---- F12 scrive sempre, e dice perche'
app.entry_search.event_generate("<KeyPress>", keysym="F12"); app.update()
dentro = open(scritto, encoding="utf-8").read()
assert dentro.count("REGISTRO DELLA TASTIERA") == 2, dentro.count("REGISTRO DELLA TASTIERA")
assert "(F12)" in dentro

# ---- e la diagnostica lo raccoglie
from inventario import diagnostica
rapporto = "\n".join(diagnostica.registro_della_tastiera([cartella]))
assert "SINTOMO" in rapporto and scritto in rapporto, rapporto[:300]
assert "nessun" in "\n".join(diagnostica.registro_della_tastiera([tempfile.mkdtemp()]))

app.destroy()
print("REGISTRO OK")

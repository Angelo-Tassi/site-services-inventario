"""Nella barra dei comandi non si perde mai un pulsante.

Tk, in una riga sola, non stringe i pulsanti che non ci stanno: semplicemente
non li dispone, e spariscono senza un errore. E' successo davvero - Reset
inventario, Ripristina, Impostazioni e Aggiorna erano fuori dallo schermo su un
monitor da 1470 px - e nessuna suite se n'era accorta. Qui si misura.
"""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import fixture
from inventario import ui
from inventario.ui import App

app = App(fixture.build())
app._initial_load()
app.update()

ATTESI = {"‹  Home", "Aggiungi", "Modifica", "Elimina", "Elimina +",
          "Sposta in stanza...", "Importa xls...", "Esporta xls...", "Stampa",
          "Aggiorna", "Impostazioni"}

def pulsanti_visibili(larghezza):
    app.geometry("%dx720" % larghezza)
    app.update(); app.update_idletasks()
    barra = app.barra
    visti, fuori = set(), []
    for gruppo in (app.gruppo_lavoro, app.gruppo_file, app.gruppo_lato):
        for c in gruppo.winfo_children():
            if c.winfo_class() != "TButton":
                continue
            testo = str(c.cget("text"))
            visti.add(testo)
            x = c.winfo_rootx() - barra.winfo_rootx()
            if (not c.winfo_ismapped() or x < -1
                    or x + c.winfo_reqwidth() > barra.winfo_width() + 1):
                fuori.append(testo)
    return visti, fuori

# ---- dalla larghezza minima della finestra in su, non manca mai niente
minima = app.wm_minsize()[0]
for larghezza in (minima, 1000, 1100, 1220, 1350, 1409, 1500, 1800):
    visti, fuori = pulsanti_visibili(larghezza)
    assert visti == ATTESI, (larghezza, sorted(ATTESI - visti))
    assert not fuori, "a %d px restano fuori: %s" % (larghezza, fuori)

# ---- larga va in una riga, stretta va a capo
pulsanti_visibili(1800)
assert app._righe_barra == "una", app._righe_barra
alta_una = app.barra.winfo_height()
pulsanti_visibili(1220)
assert app._righe_barra == "due", app._righe_barra
assert app.barra.winfo_height() > alta_una, "la seconda riga deve occupare spazio"
pulsanti_visibili(minima)
assert app._righe_barra == "due_basso", app._righe_barra

# ---- e si torna indietro: allargando si ricompone in una riga sola
pulsanti_visibili(1800)
assert app._righe_barra == "una"
assert app.barra.winfo_height() == alta_una
assert app.sep_barra.winfo_ismapped(), "il separatore torna fra i due gruppi"

# ---- i comandi che riscrivono l'inventario non stanno piu' nella barra
nella_barra = set()
def gira(w):
    if w.winfo_class() == "TButton":
        nella_barra.add(str(w.cget("text")))
    for c in w.winfo_children():
        gira(c)
gira(app.barra)
for spostato in ("Reset inventario", "Ripristina", "Salva copia in locale..."):
    assert spostato not in nella_barra, spostato
assert "Impostazioni" in nella_barra, "la porta per raggiungerli deve restare"

app.destroy()
print("BARRA COMANDI OK")

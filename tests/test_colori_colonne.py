"""Ogni colonna ha un colore suo, per non perderla scorrendo un elenco largo.

ttk non sa colorare una cella ne' una colonna: lo stile della tabella vale per
tutte. Si colora quello che si puo' - la riga verticale che divide due colonne e
una barretta nell'intestazione - lasciando il testo scuro su bianco, che e' la
combinazione piu' leggibile che ci sia.
"""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import fixture
from inventario import theme
from inventario.ui import App

def luminanza(colore):
    r, g, b = (int(colore[i:i + 2], 16) / 255 for i in (1, 3, 5))
    def c(x):
        return x / 12.92 if x <= 0.03928 else ((x + 0.055) / 1.055) ** 2.4
    r, g, b = c(r), c(g), c(b)
    return 0.2126 * r + 0.7152 * g + 0.0722 * b

def contrasto(a, b):
    chiaro, scuro = sorted((luminanza(a), luminanza(b)), reverse=True)
    return (chiaro + 0.05) / (scuro + 0.05)

# ---- i colori si vedono sul bianco della tabella
for campo, colore in theme.COLORE_COLONNA.items():
    rapporto = contrasto(colore, theme.CARD)
    assert rapporto >= 3.0, "%s: %s su bianco fa %.1f:1" % (campo, colore, rapporto)
assert contrasto(theme.COLORE_COLONNA_ALTRO, theme.CARD) >= 3.0

# ---- campi vicini per significato condividono il colore, gli altri no
assert theme.COLORE_COLONNA["prestato_a"] == theme.COLORE_COLONNA["prestato_il"]
assert theme.COLORE_COLONNA["stato"] != theme.COLORE_COLONNA["stanza"]
assert theme.COLORE_COLONNA["asset_tag"] != theme.COLORE_COLONNA["note"]

app = App(fixture.build())
app._initial_load()
app.geometry("1250x640")
app.update()

def confini(app):
    """(indice della colonna, riga verticale) per i confini che si vedono.

    Le colonne si adattano al contenuto, quindi la tabella puo' essere piu'
    larga della finestra: i confini oltre il bordo non si disegnano, ed e'
    giusto cosi'.
    """
    colonne = app._columns()
    larghezze = [int(app.tree.column(c, "width")) for c in colonne]
    disponibile = app.tree.winfo_width()
    # le coordinate di Tk partono dal bordo esterno della tabella, il conto
    # delle larghezze da zero: senza il rientro il confine atteso cade due
    # pixel a sinistra di quello vero
    fuori, x = [], app._rientro_tabella()
    for indice in range(len(colonne) - 1):
        x += larghezze[indice]
        if 0 < x < disponibile:
            fuori.append((indice, x, app._righelli[indice]))
    return fuori

app._sync_righelli()
app.update()
colonne = app._columns()
visibili = confini(app)
assert visibili, "nessun confine visibile: la tabella non e' stata disegnata"
disegnati = [r for r in app._righelli if r.winfo_ismapped()]
assert len(disegnati) == len(visibili), (len(disegnati), len(visibili))

# ---- ogni riga sta al confine giusto e porta il colore della colonna che apre
for indice, x, riga in visibili:
    assert riga.winfo_ismapped(), indice
    assert abs(riga.winfo_x() - (x - 1)) <= 1, (indice, riga.winfo_x(), x)
    atteso = theme.COLORE_COLONNA.get(colonne[indice + 1], theme.COLORE_COLONNA_ALTRO)
    assert str(riga.cget("bg")) == atteso, (colonne[indice + 1], riga.cget("bg"))
    assert riga.winfo_height() == app.tree.winfo_height()

# ---- e ogni colonna di dati porta la sua barretta nell'intestazione
# (la casella di selezione e la colonna dei pulsanti non hanno intestazione)
for campo in app._campi_visibili():
    assert app.tree.heading(campo)["image"], campo
for campo in ("asset_tag", "stato"):
    assert app._segno_colonna(campo) is app._segno_colonna(campo), "riusata, non ricreata"

# ---- entrando in una stanza le colonne cambiano, e le righe le seguono
app.show_room(fixture.KIOSK)
app.update()
app._sync_righelli()
app.update()
colonne = app._columns()
visibili = confini(app)
assert visibili, "nel Kiosk nessun confine visibile"
assert len([r for r in app._righelli if r.winfo_ismapped()]) == len(visibili)
for indice, _x, riga in visibili:
    atteso = theme.COLORE_COLONNA.get(colonne[indice + 1], theme.COLORE_COLONNA_ALTRO)
    assert str(riga.cget("bg")) == atteso, (colonne[indice + 1], riga.cget("bg"))

app.destroy()
print("COLORI COLONNE OK")

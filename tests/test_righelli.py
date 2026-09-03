"""Le righe verticali colorate stanno sul bordo delle colonne, anche scorrendo.

Servono a non perdere la colonna in un elenco largo, quindi devono cadere
esattamente fra una colonna e l'altra. Scorrendo verso destra finivano invece in
mezzo al testo: il conto dello scostamento trattava a parte l'estremo destro e
li' lo azzerava, come se la tabella non fosse stata scorsa affatto.

Qui la posizione calcolata si confronta con quella vera, chiesta a Tk con
bbox(): e' l'unico modo di accorgersene senza guardare lo schermo.
"""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import fixture
from inventario.ui import App

app = App(fixture.build())
app._initial_load()
app.geometry("980x600")
app.update(); app.update_idletasks()

def scarti(dove):
    """Di quanti pixel ogni righello sbaglia il bordo della sua colonna."""
    tree = app.tree
    tree.xview_moveto(dove)
    app.update(); app.update_idletasks()
    colonne = app._columns()
    riga = tree.get_children()[0]
    fuori = []
    for indice, colonna in enumerate(colonne[:-1]):
        riquadro = tree.bbox(riga, colonna)
        if not riquadro:
            continue
        bordo = riquadro[0] + riquadro[2]
        if not (0 < bordo < tree.winfo_width()):
            continue                      # colonna fuori dalla vista: niente riga
        righello = app._righelli[indice]
        if not righello.winfo_ismapped():
            fuori.append((colonna, "non disegnato"))
            continue
        # il righello e' largo 2 px e sta a cavallo del bordo
        fuori.append((colonna, righello.winfo_x() + 1 - bordo))
    return fuori

# ---- dentro una stanza, dove le colonne superano la finestra
app.show_room(fixture.BAU)
app.update(); app.update_idletasks()
assert app.tree.xview()[1] < 1.0, "la prova ha senso solo se si puo' scorrere"

for dove in (0.0, 0.1, 0.25, 0.5, 0.75, 1.0):
    misure = scarti(dove)
    assert misure, ("nessun righello a xview %s" % dove)
    for colonna, scarto in misure:
        assert scarto == 0, "a xview %.2f la riga di %s sbaglia di %s px" % (
            dove, colonna, scarto)

# ---- il caso che si vedeva a occhio: tutto a destra
app.tree.xview_moveto(1.0)
app.update(); app.update_idletasks()
primo = app.tree.xview()[0]
assert primo > 0, "la tabella deve essere davvero scorsa"
# il righello della prima colonna ancora visibile non puo' stare dove starebbe
# a tabella ferma: sarebbe il difetto di prima
larghezze = [int(app.tree.column(c, "width")) for c in app._columns()]
assert app._righelli[0].winfo_x() != larghezze[0] - 1 or not \
    app._righelli[0].winfo_ismapped(), "lo scostamento e' stato ignorato"

# ---- e in panoramica, dove le colonne ci stanno tutte
app.show_home()
app.update(); app.update_idletasks()
for colonna, scarto in scarti(0.0):
    assert scarto == 0, "in home la riga di %s sbaglia di %s px" % (colonna, scarto)

# ---- un elenco vuoto non fa saltare il disegno
app.var_search.set("nessun-dispositivo-con-questo-nome")
app.update(); app.update_idletasks()
app._sync_righelli()
app.var_search.set("")
app.update(); app.update_idletasks()

app.destroy()
print("RIGHELLI OK")

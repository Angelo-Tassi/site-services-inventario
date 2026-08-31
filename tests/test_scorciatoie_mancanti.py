"""La tabella compare anche dove una scorciatoia non esiste.

Il caso vero: su Windows i pulsanti 6 e 7 del mouse - la rotella orizzontale di
X11 - non esistono, e Tk rifiuta il collegamento con un errore. Quell'errore
interrompeva la costruzione della tabella a meta': i dispositivi c'erano, le
schede delle stanze li contavano, ma l'elenco restava bianco. Nessun sintomo
rimandava a una scorciatoia da tastiera.
"""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import fixture
import tkinter as tk
from inventario.ui import App

app = App(fixture.build())
app._initial_load()
assert len(app.tree.get_children()) == 13, "la tabella parte piena"

# ---- si finge il rifiuto di Windows su ogni collegamento non universale
rifiutati = []
vero_bind = tk.Widget.bind

def bind_alla_windows(self, sequence=None, func=None, add=None):
    if sequence in ("<Button-6>", "<Button-7>"):
        rifiutati.append(sequence)
        raise tk.TclError('bad button number "%s"' % sequence[-2])
    return vero_bind(self, sequence, func, add)

tk.Widget.bind = bind_alla_windows
try:
    app.show_room(fixture.KIOSK)     # ricostruisce la tabella
    # tutte e due vengono tentate: il rifiuto della prima non ferma la seconda
    assert rifiutati == ["<Button-6>", "<Button-7>"], rifiutati
    assert app.tree is not None
    assert len(app.tree.get_children()) == 5, len(app.tree.get_children())
    assert app.tree.winfo_exists()
    # e il resto della tabella e' costruito per intero
    assert str(app.tree.cget("xscrollcommand")), "lo scorrimento orizzontale c'e' lo stesso"
    barre = [w for w in app.tree.master.winfo_children()
             if w.winfo_class() == "TScrollbar"]
    assert len(barre) == 2, len(barre)
    app.show_home()
    assert len(app.tree.get_children()) == 13
finally:
    tk.Widget.bind = vero_bind

app.destroy()
print("SCORCIATOIE MANCANTI OK")

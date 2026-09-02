"""Copia e incolla in ogni campo, e dall'elenco verso Excel.

Le scorciatoie di Tk non sono le stesse su tutti i sistemi ne' con tutte le
disposizioni di tastiera, e in qualche caso non funzionano affatto: qui si
verifica che il programma le colleghi da se', in tutte le classi di campo, e che
dall'elenco si porti via un identificativo pronto da incollare.
"""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import fixture
import tkinter as tk
from tkinter import ttk
from inventario.ui import ACTION_COLUMN, CHECK_COLUMN, App

class Finto(object):
    """Un evento con il solo campo che serve ai gestori."""
    def __init__(self, widget):
        self.widget = widget

app = App(fixture.build())
app._initial_load()
app.update()
appunti = app._appunti

# ---- i collegamenti ci sono su tutte le classi di campo, in tutte le forme
for classe in ("Entry", "TEntry", "TCombobox", "Text"):
    for sequenza in ("<Control-c>", "<Control-v>", "<Control-x>", "<Control-a>",
                     "<Command-c>", "<Command-v>", "<<Copy>>", "<<Paste>>",
                     "<<Cut>>", "<Button-3>"):
        assert app.bind_class(classe, sequenza), (classe, sequenza)

# ---- si incolla in un campo di testo
app.clipboard_clear(); app.clipboard_append("IT-INCOLLATO")
campo = ttk.Entry(app); campo.pack(); app.update()
appunti["incolla"](Finto(campo))
assert campo.get() == "IT-INCOLLATO", campo.get()

# ---- e si copia da un campo di testo
app.clipboard_clear(); app.clipboard_append("altro")
appunti["seleziona_tutto"](Finto(campo))
appunti["copia"](Finto(campo))
assert app.clipboard_get() == "IT-INCOLLATO", app.clipboard_get()

# ---- taglia svuota il campo ma lascia il testo negli appunti
appunti["taglia"](Finto(campo))
assert campo.get() == "" and app.clipboard_get() == "IT-INCOLLATO"

# ---- una tendina scrivibile accetta l'incollaggio, come quella del tipo
tendina = ttk.Combobox(app, values=["Laptop", "Tablet"]); tendina.pack(); app.update()
appunti["incolla"](Finto(tendina))
assert tendina.get() == "IT-INCOLLATO", tendina.get()

# ---- una in sola lettura non si scrive, ma si copia
sola_lettura = ttk.Combobox(app, values=["Disponibile"], state="readonly")
sola_lettura.set("Disponibile"); sola_lettura.pack(); app.update()
appunti["incolla"](Finto(sola_lettura))
assert sola_lettura.get() == "Disponibile", "in sola lettura non si scrive"
app.clipboard_clear(); app.clipboard_append("x")
appunti["copia"](Finto(sola_lettura))
assert app.clipboard_get() == "Disponibile", app.clipboard_get()

# ---- un'area di testo prende anche piu' righe: e' il caso di Elimina +
area = tk.Text(app, height=3); area.pack(); app.update()
app.clipboard_clear(); app.clipboard_append("IT-0101\nIT-0102\nIT-0103")
appunti["incolla"](Finto(area))
assert area.get("1.0", "end").split() == ["IT-0101", "IT-0102", "IT-0103"]

# ---- dall'elenco si copia la riga selezionata, incolonnata per Excel
app.show_home(); app.update()
app.tree.selection_set("IT-0101")
app.clipboard_clear(); app.clipboard_append("vuoto")
app._copia_selezione()
copiata = app.clipboard_get()
colonne = [c for c in app._columns() if c not in (CHECK_COLUMN, ACTION_COLUMN)]
assert copiata.count("\t") == len(colonne) - 1, copiata
assert copiata.startswith("IT-0101\t"), copiata

# ---- oppure il solo identificativo, che e' quello che serve a Elimina +
app._copia_selezione(solo_identificativo=True)
assert app.clipboard_get() == "IT-0101", app.clipboard_get()

# ---- senza selezione non si copia niente e lo si dice
app.tree.selection_remove(*app.tree.selection())
for tag in app.tree.get_children():
    app.tree.set(tag, CHECK_COLUMN, "")
app.clipboard_clear(); app.clipboard_append("intatto")
app._copia_selezione()
assert app.clipboard_get() == "intatto"
assert "nessuna riga" in app.var_status.get().lower(), app.var_status.get()

app.destroy()
print("APPUNTI OK")

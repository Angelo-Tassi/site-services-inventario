"""Nessuna suite deve restare ferma ad aspettare un clic.

Un messagebox aperto durante i test blocca il processo finche' qualcuno non lo
chiude a mano: run_all.py lo uccide dopo due minuti e scrive BLOCCATA, senza
dire su che cosa. E' successo davvero, e non tutte le volte - dipendeva da chi
era davanti allo schermo in quel momento.

fixture risponde da sola a ogni finestra. Qui si verifica che risponda come
serve: No dove la risposta avvierebbe qualcosa, Ok dove serve solo confermare.
"""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import fixture
from tkinter import filedialog, messagebox

# ---- le risposte che non fanno partire niente
assert messagebox.askyesno("Conviene una copia", "?") is False, \
    "il promemoria della copia locale deve essere rifiutato, o parte un salvataggio"
assert messagebox.askyesnocancel("x", "?") is False
assert messagebox.askokcancel("Conferma", "?") is True, \
    "una conferma va accettata, o meta' delle prove non arriva in fondo"
assert messagebox.showinfo("x", "y") == "ok"
assert messagebox.showwarning("x", "y") == "ok"
assert messagebox.showerror("x", "y") == "ok"
assert filedialog.asksaveasfilename() == "", "la scelta del file va annullata"
assert filedialog.askdirectory() == ""

# ---- e ogni domanda resta scritta, cosi' una suite puo' controllarla
titoli = [t for _tipo, t, _m in fixture.popup]
assert "Conviene una copia" in titoli, fixture.popup
assert ("askyesno", "Conviene una copia", "?") in fixture.popup

# ---- le finestre del programma non aspettano per sempre
from inventario import ui
assert getattr(ui._Modal, "_show_sorvegliato", False), \
    "senza la scadenza una finestra dimenticata ferma la suite per due minuti"

# ---- la rete non impedisce a una suite di dare le sue risposte
fixture.popup.clear()
messagebox.askyesno = lambda t, m, **k: True
assert messagebox.askyesno("x", "?") is True, "l'ultima assegnazione deve vincere"
fixture.silenzia_i_popup()
assert messagebox.askyesno("x", "?") is False, "e si puo' rimettere la rete"

print("POPUP ZITTI OK")

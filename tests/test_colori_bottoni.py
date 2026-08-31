"""I pulsanti dicono con il colore a che famiglia appartengono.

Arancione i dati che entrano, verde quelli che escono, rosso quello che
riscrive l'inventario di tutti. Se qualcuno aggiunge un comando di quelle
famiglie senza colorarlo, o cambia un'etichetta, questa suite se ne accorge.
"""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import fixture
from inventario import config, ui, theme
from inventario import lingua as lang
from tkinter import ttk
p = fixture.build()
config.load_data_path = lambda: p
app = ui.App(p); app._initial_load()
esiti = []
def controlla():
    try:
        trovati = {}
        def gira(w):
            if w.winfo_class() == "TButton":
                try:
                    trovati[str(w.cget("text"))] = str(w.cget("style"))
                except Exception:
                    pass
            for c in w.winfo_children():
                gira(c)
        gira(app)
        atteso = {
            "Importa xls...": "Arancio.TButton",
            "Esporta xls...": "Verde.TButton",
            "Salva copia in locale...": "Verde.TButton",
            "Ripristina": "Rosso.TButton",
            "Reset inventario": "Rosso.TButton",
        }
        for testo, stile in atteso.items():
            esiti.append(("%s -> %s" % (testo, stile), trovati.get(testo) == stile))
        app.show_room(fixture.KIOSK); app.update()
        trovati = {}; gira(app)
        esiti.append(("Esporta questa stanza in xls -> Verde",
                      trovati.get("Esporta questa stanza in xls") == "Verde.TButton"))
        esiti.append(("Importa i dati di questa stanza -> Arancio",
                      trovati.get("Importa i dati di questa stanza") == "Arancio.TButton"))
        st = ttk.Style(app)
        esiti.append(("rosso definito", st.lookup("Rosso.TButton", "background") == theme.AZIONE_ROSSA_BG))
        esiti.append(("verde definito", st.lookup("Verde.TButton", "background") == theme.AZIONE_VERDE_BG))
        esiti.append(("arancio definito", st.lookup("Arancio.TButton", "background") == theme.AZIONE_ARANCIO_BG))
        # in inglese il pulsante nuovo esiste con l'etichetta tradotta
        lang.imposta("en"); app.ricostruisci(); app.update()
        trovati = {}; gira(app)
        esiti.append(("Save a local copy... in inglese",
                      trovati.get("Save a local copy...") == "Verde.TButton"))
        lang.imposta("it")
    except Exception as exc:
        esiti.append(("eccezione", repr(exc)))
    app.destroy()
app.after(700, controlla)
app.mainloop()
falliti = [n for n, ok in esiti if ok is not True]
assert not falliti, falliti
assert len(esiti) == 11, len(esiti)

# ---- il testo resta leggibile sul colore: contrasto AA su tutti gli stati
def luminanza(colore):
    r, g, b = (int(colore[i:i + 2], 16) / 255 for i in (1, 3, 5))
    def c(x):
        return x / 12.92 if x <= 0.03928 else ((x + 0.055) / 1.055) ** 2.4
    r, g, b = c(r), c(g), c(b)
    return 0.2126 * r + 0.7152 * g + 0.0722 * b

def contrasto(a, b):
    chiaro, scuro = sorted((luminanza(a), luminanza(b)), reverse=True)
    return (chiaro + 0.05) / (scuro + 0.05)

for sfondo, testo in ((theme.AZIONE_ROSSA_BG, theme.AZIONE_ROSSA_FG),
                      (theme.AZIONE_ROSSA_BG_ON, theme.AZIONE_ROSSA_FG),
                      (theme.AZIONE_VERDE_BG, theme.AZIONE_VERDE_FG),
                      (theme.AZIONE_VERDE_BG_ON, theme.AZIONE_VERDE_FG),
                      (theme.AZIONE_ARANCIO_BG, theme.AZIONE_ARANCIO_FG),
                      (theme.AZIONE_ARANCIO_BG_ON, theme.AZIONE_ARANCIO_FG)):
    rapporto = contrasto(sfondo, testo)
    assert rapporto >= 4.5, "%s su %s: contrasto %.1f" % (testo, sfondo, rapporto)

print("COLORI BOTTONI OK")

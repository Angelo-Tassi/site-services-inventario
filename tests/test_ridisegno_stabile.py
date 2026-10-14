"""A finestra ferma sync non modifica widget; lo schermo piccolo contiene l'avvio."""
import os
import sys
import time
from contextlib import ExitStack
from unittest.mock import patch

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import fixture
from inventario.ui import App


def respira(app, durata=0.2):
    fine = time.monotonic() + durata
    while time.monotonic() < fine:
        app.update()
        time.sleep(0.005)


with patch.object(App, "winfo_screenwidth", return_value=1280), \
        patch.object(App, "winfo_screenheight", return_value=720):
    app = App(fixture.build())
    try:
        app._initial_load()
        respira(app)
        if app.tk.call("tk", "windowingsystem") == "win32":
            # su Windows lo schermo piccolo fa partire la finestra
            # massimizzata: la misura e' quella dello schermo vero del PC, non
            # di quello simulato
            assert app.state() == "zoomed", app.state()
        else:
            assert app.winfo_width() <= 1280 - 60
            assert app.winfo_height() <= 720 - 100
        minimo = app.minsize()
        assert minimo[0] <= 1280 - 60 and minimo[1] <= 720 - 100, minimo
        app.show_room(fixture.KIOSK)
        respira(app)
        assert app._row_buttons, "la prova deve misurare anche i pulsanti"
        # Tutti i widget sovrapposti sono sorvegliati: anche configure senza
        # cambiamenti e place sulle stesse coordinate causano lavoro a Windows.
        with ExitStack() as contesto:
            controlli = []
            for widget in list(app._righelli) + list(app._row_buttons.values()):
                for metodo in ("configure", "place", "place_forget", "destroy"):
                    controlli.append(contesto.enter_context(patch.object(
                        widget, metodo, wraps=getattr(widget, metodo))))
            prima = dict(app._ridisegni)
            for _ in range(20):
                app._sync_righelli()
                app._sync_row_buttons()
            app.update()
            assert app._ridisegni["righelli"] - prima["righelli"] == 1
            assert app._ridisegni["pulsanti"] - prima["pulsanti"] == 1
            assert all(c.call_count == 0 for c in controlli), [c for c in controlli if c.call_count]
            prima = dict(app._ridisegni)
            respira(app)
            assert app._ridisegni == prima, (prima, app._ridisegni)
        # Lo scorrimento deve continuare a spostare davvero righelli/pulsanti.
        app.tree.xview_moveto(1)
        app.update()
        app.tree.xview_moveto(0)
        app.update()
        # Vicino al punto in cui compare il separatore, il layout converge
        # sia partendo largo sia partendo stretto.
        soglia = (app.gruppo_lavoro.winfo_reqwidth() + app.gruppo_lato.winfo_reqwidth()
                  + sum(c.winfo_reqwidth() + 6 for c in app.gruppo_file.winfo_children()
                        if c is not app.sep_barra) + app.sep_barra.winfo_reqwidth() + 20 + 32)
        for delta in (-20, -1, 0, 1, 10, 20, -1):
            app.geometry("%dx620" % (soglia + delta))
            app.update()
            prima = app._righe_barra
            for _ in range(5):
                app._sistema_barra()
                app.update()
                assert app._righe_barra == prima
        # Una sync in coda non usa i figli della tabella distrutta.
        app._sync_row_buttons()
        app.show_home()
        app.update()
    finally:
        app.destroy()


class SchermoWindows:
    def __init__(self):
        self.tk = self
        self.stato = None

    def call(self, *args):
        return "win32"

    def winfo_screenwidth(self):
        return 1280

    def winfo_screenheight(self):
        return 720

    def geometry(self, valore):
        self.dimensione = tuple(map(int, valore.split("x")))

    def minsize(self, *valore):
        self.minimo = valore

    def state(self, valore):
        self.stato = valore


schermo = SchermoWindows()
App._geometria_iniziale(schermo)
assert schermo.dimensione == (1220, 620)
assert schermo.minimo == (980, 560)
assert schermo.stato == "zoomed"
print("RIDISEGNO STABILE OK")

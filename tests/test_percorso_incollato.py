"""Il percorso della share si accetta in tutte le forme in cui Windows lo copia.

Chi installa non deve ripulire a mano quello che ha incollato: le virgolette di
"Copia come percorso", le barre al contrario della barra degli indirizzi, un
file:// o una barra di troppo sono tutte la stessa cartella.
"""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import fixture
from inventario import config, configura

# ---- come si comporta su Windows, dove le barre sono rovesciate
vero_sep, vero_join = os.sep, os.path.join
os.sep = "\\"
os.path.join = lambda *p: "\\".join(x.rstrip("\\") for x in p if x)
try:
    atteso = "F:\\Inventario"
    for scritto in ('"F:\\Inventario"', "F:/Inventario", "F://inventario".replace(
                        "inventario", "Inventario"), "F:\\Inventario\\",
                    "  F:\\Inventario  ", "'F:\\Inventario'",
                    "file:///F:/Inventario"):
        assert configura.normalizza(scritto) == atteso, (scritto,
                                                         configura.normalizza(scritto))

    # il doppio backslash iniziale di un percorso di rete non va perso
    rete = "\\\\server\\Condivisa\\Inventario"
    for scritto in (rete, '"' + rete + '"', "//server/Condivisa/Inventario",
                    rete + "\\", "\\\\\\server\\Condivisa\\Inventario"):
        assert configura.normalizza(scritto) == rete, (scritto,
                                                       configura.normalizza(scritto))

    # ---- e da qualunque forma si arriva sempre allo stesso file
    atteso_file = "F:\\Inventario\\Produzione\\Inventario.xlsx"
    for scritto in ('"F:\\Inventario"', "F:/Inventario",
                    "F:\\Inventario\\Produzione",
                    "F:\\Inventario\\Produzione\\Inventario.xlsx"):
        assert configura.percorso_inventario(scritto) == atteso_file, scritto
finally:
    os.sep, os.path.join = vero_sep, vero_join

# ---- su questo sistema le forme normali restano valide
import tempfile
base = tempfile.mkdtemp()
atteso = os.path.join(base, config.NOME_PRODUZIONE, config.DATA_FILE_NAME)
for scritto in (base, '"%s"' % base, base + os.sep):
    assert configura.percorso_inventario(scritto) == atteso, scritto

print("PERCORSO INCOLLATO OK")

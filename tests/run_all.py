#!/usr/bin/env python3
"""Esegue tutte le suite di test. Uso:  .venv/bin/python tests/run_all.py"""

import os
import subprocess
import sys

QUI = os.path.dirname(os.path.abspath(__file__))
# Sul runner Windows di GitHub Tk e' piu' lento a disegnare: piu' respiro
TIMEOUT = 240 if sys.platform == "win32" else 120
# Suite che hanno bisogno di cose che su Windows non ci sono (permessi POSIX,
# lock di file alla Unix): si dichiarano qui, e la riga dice perche'
SOLO_POSIX = {
    "test_config.py": "rende una cartella di sola lettura con chmod, che su "
                      "Windows non ha effetto sulle cartelle",
}


def annota_su_actions(nome, coda):
    """Su GitHub Actions, il fallimento diventa un'annotazione del job.

    Le annotazioni si leggono dall'API pubblica anche senza essere loggati,
    i log completi no: e' l'unico modo per vedere da un Mac senza account che
    cosa e' fallito sul runner Windows. Il testo va su una riga sola, con i
    ritorni a capo codificati come vuole il formato.
    """
    if not os.environ.get("GITHUB_ACTIONS"):
        return
    messaggio = coda[-3000:].replace("%", "%25").replace("\r", "%0D") \
        .replace("\n", "%0A")
    print("::error title=%s::%s" % (nome, messaggio))
    sys.stdout.flush()


def main():
    suite = sorted(f for f in os.listdir(QUI)
                   if f.startswith("test_") and f.endswith(".py"))
    falliti = []
    saltate = 0
    for nome in suite:
        if sys.platform == "win32" and nome in SOLO_POSIX:
            saltate += 1
            print("%-26s saltata su Windows: %s" % (nome, SOLO_POSIX[nome]))
            continue
        # Con un limite di tempo: una finestra di dialogo lasciata aperta per
        # sbaglio bloccherebbe l'intera esecuzione senza dire perche'.
        try:
            esito = subprocess.run([sys.executable, os.path.join(QUI, nome)],
                                   capture_output=True, text=True, timeout=TIMEOUT)
        except subprocess.TimeoutExpired:
            falliti.append(nome)
            print("%-26s BLOCCATA dopo %d secondi" % (nome, TIMEOUT))
            annota_su_actions(nome, "bloccata dopo %d secondi: probabile finestra "
                                    "di dialogo in attesa" % TIMEOUT)
            print("    probabile finestra di dialogo in attesa: sostituisci "
                  "messagebox nel test")
            continue
        ultima = (esito.stdout.strip().splitlines() or ["(nessun output)"])[-1]
        if esito.returncode == 0:
            print("%-26s %s" % (nome, ultima))
        else:
            falliti.append(nome)
            print("%-26s FALLITO" % nome)
            coda = (esito.stdout + esito.stderr).strip()[-1200:]
            print(coda)
            annota_su_actions(nome, coda)
    print()
    print("%d suite, %d fallite%s" % (len(suite), len(falliti),
                                     ", %d saltate" % saltate if saltate else ""))
    return 1 if falliti else 0


if __name__ == "__main__":
    sys.exit(main())

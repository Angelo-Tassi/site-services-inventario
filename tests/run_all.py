#!/usr/bin/env python3
"""Esegue tutte le suite di test. Uso:  .venv/bin/python tests/run_all.py"""

import os
import subprocess
import sys

QUI = os.path.dirname(os.path.abspath(__file__))
TIMEOUT = 120


def main():
    suite = sorted(f for f in os.listdir(QUI)
                   if f.startswith("test_") and f.endswith(".py"))
    falliti = []
    for nome in suite:
        # Con un limite di tempo: una finestra di dialogo lasciata aperta per
        # sbaglio bloccherebbe l'intera esecuzione senza dire perche'.
        try:
            esito = subprocess.run([sys.executable, os.path.join(QUI, nome)],
                                   capture_output=True, text=True, timeout=TIMEOUT)
        except subprocess.TimeoutExpired:
            falliti.append(nome)
            print("%-26s BLOCCATA dopo %d secondi" % (nome, TIMEOUT))
            print("    probabile finestra di dialogo in attesa: sostituisci "
                  "messagebox nel test")
            continue
        ultima = (esito.stdout.strip().splitlines() or ["(nessun output)"])[-1]
        if esito.returncode == 0:
            print("%-26s %s" % (nome, ultima))
        else:
            falliti.append(nome)
            print("%-26s FALLITO" % nome)
            print((esito.stdout + esito.stderr).strip()[-1200:])
    print()
    print("%d suite, %d fallite" % (len(suite), len(falliti)))
    return 1 if falliti else 0


if __name__ == "__main__":
    sys.exit(main())

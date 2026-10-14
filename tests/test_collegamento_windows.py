"""Collegamento desktop: pacchetto completo, percorsi difficili ed errori reali.

Su Windows si eseguono il .bat e il .ps1 distribuiti, creando veri .lnk in
una scrivania temporanea. Sugli altri sistemi si verificano file e codifica.
"""
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile

RADICE = Path(__file__).resolve().parent.parent
BAT = 'Crea collegamento sul desktop.bat'
PS1 = 'Crea collegamento sul desktop.ps1'
AVVIO = 'Avvia Inventario.bat'

for nome in (BAT, AVVIO):
    dati = (RADICE / nome).read_bytes()
    assert b'\r\n' in dati and b'\n' not in dati.replace(b'\r\n', b''), nome
    dati.decode('ascii')
assert (RADICE / PS1).read_bytes().startswith(b'\xef\xbb\xbf'), 'BOM per PowerShell 5.1'

# Sul runner la suite gira dentro ProvaSuite, senza il sorgente del workflow.
workflow = RADICE / '.github' / 'workflows' / 'build-windows.yml'
if workflow.exists():
    testo = workflow.read_text(encoding='utf-8')
    prima_suite = testo.split('- name: Suite di test su Windows')[0]
    for nome in (BAT, PS1, AVVIO):
        assert 'Copy-Item "%s" Pacchetto\\' % nome in prima_suite, nome

if sys.platform != 'win32':
    print('COLLEGAMENTO WINDOWS OK (file e codifica; esecuzione COM solo su Windows)')
    sys.exit(0)

powershell = str(Path(os.environ['SystemRoot']) / 'System32' / 'WindowsPowerShell' / 'v1.0' / 'powershell.exe')
ambiente = dict(os.environ, INV_SENZA_PAUSA='1')


def esegui_bat(cartella, desktop):
    # cmd richiede le virgolette esterne oltre a quelle dei singoli percorsi.
    comando = '"%s" /d /s /c ""%s" -Desktop "%s""' % (
        os.environ['COMSPEC'], cartella / BAT, desktop)
    return subprocess.run(comando, env=ambiente, capture_output=True, timeout=30)


def leggi_collegamento(percorso):
    codice = """
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)
$s = (New-Object -ComObject WScript.Shell).CreateShortcut($env:INV_PROVA_COLLEGAMENTO)
[pscustomobject]@{Target=$s.TargetPath; Cartella=$s.WorkingDirectory;
                 Icona=$s.IconLocation; Stato=$s.WindowStyle} | ConvertTo-Json -Compress
"""
    esito = subprocess.run([powershell, '-NoProfile', '-NonInteractive', '-Command', codice],
                           env=dict(ambiente, INV_PROVA_COLLEGAMENTO=str(percorso)),
                           capture_output=True, timeout=30)
    assert esito.returncode == 0, esito.stderr
    return json.loads(esito.stdout)


with tempfile.TemporaryDirectory() as temporanea:
    # Spazi, apostrofi, accenti, parentesi, !, & e [] devono restare dati.
    pacchetto = Path(temporanea) / "Inventario di D'Angelo è pronto! & (prove) [1]"
    desktop = Path(temporanea) / "OneDrive - Società D'Angelo" / 'Desktop [ufficio]'
    pacchetto.mkdir()
    desktop.mkdir(parents=True)
    for nome in (BAT, PS1, AVVIO):
        shutil.copy2(RADICE / nome, pacchetto / nome)

    esito = esegui_bat(pacchetto, desktop)
    assert esito.returncode == 0, (esito.stdout, esito.stderr)
    collegamento = desktop / 'Inventario dispositivi.lnk'
    assert collegamento.exists() and (pacchetto / collegamento.name).exists()
    dati = leggi_collegamento(collegamento)
    assert Path(dati['Target']) == pacchetto / AVVIO, dati
    assert Path(dati['Cartella']) == pacchetto, dati
    assert 'shell32.dll' in dati['Icona'].lower(), dati
    assert dati['Stato'] == 7, dati
    assert (pacchetto / '.sbloccato').exists()

    # Con il Python incorporabile l'icona viene da pythonw.exe; il secondo
    # avvio sovrascrive il collegamento, senza ripetere lo sblocco.
    (pacchetto / 'python').mkdir()
    shutil.copy2(powershell, pacchetto / 'python' / 'pythonw.exe')
    prima = (pacchetto / '.sbloccato').stat().st_mtime_ns
    esito = esegui_bat(pacchetto, desktop)
    assert esito.returncode == 0, (esito.stdout, esito.stderr)
    dati = leggi_collegamento(collegamento)
    assert 'pythonw.exe' in dati['Icona'].lower(), dati
    assert (pacchetto / '.sbloccato').stat().st_mtime_ns == prima

    # Un desktop inesistente deve fallire, senza creare una cartella diversa.
    esito = esegui_bat(pacchetto, desktop / 'inesistente')
    assert esito.returncode != 0, esito.stdout
    assert not (desktop / 'inesistente').exists()

    # La cartella di distribuzione puo' non essere scrivibile: la copia locale
    # e' facoltativa e un .lnk sul desktop riuscito resta un successo.
    (pacchetto / collegamento.name).unlink()
    (pacchetto / collegamento.name).mkdir()
    esito = esegui_bat(pacchetto, desktop)
    assert esito.returncode == 0 and collegamento.exists(), (esito.stdout, esito.stderr)

    for nome in (PS1, AVVIO):
        mancante = pacchetto / nome
        copia = mancante.read_bytes()
        mancante.unlink()
        collegamento.unlink()
        esito = esegui_bat(pacchetto, desktop)
        assert esito.returncode != 0, (nome, esito.stdout)
        assert not collegamento.exists(), nome
        mancante.write_bytes(copia)

print('COLLEGAMENTO WINDOWS OK (bat, COM, percorsi, OneDrive, icone ed errori)')

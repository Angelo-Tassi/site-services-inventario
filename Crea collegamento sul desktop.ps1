# Crea il collegamento con Windows PowerShell 5.1, senza dipendenze aggiuntive.
# -Desktop permette anche la verifica in una cartella temporanea, senza
# toccare la scrivania di chi esegue i test. Nell'uso normale decide Windows.
[CmdletBinding()]
param([string]$Desktop)

$ErrorActionPreference = 'Stop'
try {
    $cartella = $PSScriptRoot
    $avvio = Join-Path $cartella 'Avvia Inventario.bat'
    if (-not (Test-Path -LiteralPath $avvio -PathType Leaf)) {
        throw 'Manca Avvia Inventario.bat: estrai il pacchetto completo.'
    }

    if (-not $Desktop) {
        $Desktop = [Environment]::GetFolderPath([Environment+SpecialFolder]::DesktopDirectory)
    }
    # Nessun ripiego su USERPROFILE\Desktop: con OneDrive potrebbe essere
    # una cartella invisibile all'utente, con un falso messaggio di successo.
    if (-not $Desktop -or -not (Test-Path -LiteralPath $Desktop -PathType Container)) {
        throw 'Desktop non trovato. Controlla la cartella Desktop in Esplora risorse.'
    }

    $python = Join-Path $cartella 'python\pythonw.exe'
    $icona = "$env:SystemRoot\System32\shell32.dll,15"
    if (Test-Path -LiteralPath $python -PathType Leaf) { $icona = "$python,0" }
    $percorso = Join-Path $Desktop 'Inventario dispositivi.lnk'
    $shell = New-Object -ComObject WScript.Shell
    $collegamento = $shell.CreateShortcut($percorso)
    $collegamento.TargetPath = $avvio
    $collegamento.WorkingDirectory = $cartella
    $collegamento.IconLocation = $icona
    $collegamento.WindowStyle = 7
    $collegamento.Description = 'Inventario iPhone, laptop e tablet'
    $collegamento.Save()
    if (-not (Test-Path -LiteralPath $percorso -PathType Leaf)) {
        throw 'Windows non ha salvato il collegamento sul desktop.'
    }
    Write-Host ''
    Write-Host 'Collegamento creato qui:'
    Write-Host "  $percorso"

    # La copia accanto al programma e' comoda, ma una share in sola lettura
    # non deve trasformare un collegamento desktop riuscito in un errore.
    $copia = Join-Path $cartella 'Inventario dispositivi.lnk'
    if ($copia -ne $percorso) {
        try {
            [IO.File]::Copy($percorso, $copia, $true)
            Write-Host "Copia disponibile anche in: $copia"
        } catch {
            Write-Warning "Collegamento desktop creato; copia nella cartella non disponibile: $($_.Exception.Message)"
        }
    }

    # Lo sblocco usa percorsi letterali; il marcatore nasce solo se riesce.
    $marcatore = Join-Path $cartella '.sbloccato'
    if (-not (Test-Path -LiteralPath $marcatore)) {
        try {
            Get-ChildItem -LiteralPath $cartella -Recurse -File | Unblock-File
            [IO.File]::WriteAllText($marcatore, '')
            $file = Get-Item -LiteralPath $marcatore -Force
            $file.Attributes = $file.Attributes -bor [IO.FileAttributes]::Hidden
        } catch {
            Write-Warning "Collegamento creato; sblocco automatico non completato: $($_.Exception.Message)"
        }
    }
    exit 0
} catch {
    Write-Host "Creazione del collegamento non riuscita: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host 'Se i criteri aziendali bloccano PowerShell, usa Invia a > Desktop su Avvia Inventario.bat.'
    exit 1
}

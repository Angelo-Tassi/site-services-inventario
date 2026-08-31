@echo off
rem ---------------------------------------------------------------------------
rem  Crea sul desktop un collegamento al programma che sta in questa cartella.
rem  Il desktop viene chiesto a Windows, non costruito a mano: con OneDrive la
rem  scrivania vera e' dentro OneDrive, e %USERPROFILE%\Desktop e' una cartella
rem  che l'utente non vede mai.
rem  Funziona con tutti e due i pacchetti: quello con Inventario.exe e quello
rem  senza, dove ad avviare il programma e' il pythonw.exe ufficiale.
rem ---------------------------------------------------------------------------
setlocal
set "INV_DIR=%~dp0"

if exist "%~dp0python\pythonw.exe" (
    set "INV_EXE=%~dp0python\pythonw.exe"
    set "INV_ARGS=""%~dp0Inventario.py"""
    set "INV_ICON=%~dp0python\pythonw.exe,0"
    if exist "%~dp0Inventario.exe" (
        echo.
        echo ATTENZIONE: in questa cartella ci sono DUE versioni del programma.
        echo   - Inventario.exe      resto di un pacchetto precedente
        echo   - python\ + Inventario.py   il pacchetto attuale
        echo Il collegamento puntera' a quello attuale. Cancella Inventario.exe
        echo e la cartella _internal: se qualcuno apre il vecchio .exe si ritrova
        echo un programma di una versione precedente, con i suoi difetti.
        echo.
    )
) else if exist "%~dp0Inventario.exe" (
    set "INV_EXE=%~dp0Inventario.exe"
    set "INV_ARGS="
    set "INV_ICON=%~dp0Inventario.exe,0"
) else (
    echo In questa cartella non c'e' ne' Inventario.exe ne' la cartella python:
    echo   %INV_DIR%
    echo Estrai qui il pacchetto per intero e rilancia questo file.
    pause& exit /b 1
)

powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$desk = [Environment]::GetFolderPath('Desktop');" ^
  "if (-not $desk -or -not (Test-Path $desk)) { $desk = Join-Path $env:USERPROFILE 'Desktop' };" ^
  "if (-not (Test-Path $desk)) { Write-Host 'Desktop non trovato.'; exit 1 };" ^
  "$lnk = Join-Path $desk 'Inventario dispositivi.lnk';" ^
  "$s = (New-Object -ComObject WScript.Shell).CreateShortcut($lnk);" ^
  "$s.TargetPath = $env:INV_EXE; $s.Arguments = $env:INV_ARGS;" ^
  "$s.WorkingDirectory = $env:INV_DIR; $s.IconLocation = $env:INV_ICON;" ^
  "$s.Description = 'Inventario laptop e tablet'; $s.Save();" ^
  "Copy-Item $lnk (Join-Path $env:INV_DIR 'Inventario dispositivi.lnk') -Force;" ^
  "Write-Host ''; Write-Host 'Collegamento creato qui:'; Write-Host ('  ' + $lnk)"
if errorlevel 1 (
    echo.
    echo Creazione del collegamento non riuscita.
    echo Puoi farlo a mano: tasto destro sul programma ^> Mostra altre opzioni
    echo ^> Invia a ^> Desktop ^(crea collegamento^).
    pause& exit /b 1
)

echo Una copia e' rimasta nella cartella di rete: gli altri utenti possono
echo trascinarla sul proprio desktop, senza eseguire nulla.
echo.
pause

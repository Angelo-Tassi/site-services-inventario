@echo off
setlocal EnableExtensions DisableDelayedExpansion
rem I percorsi restano dati: niente codice PowerShell composto con apici o ^.
rem Lo script accanto usa il desktop indicato da Windows, anche con OneDrive.
set "INV_DIR=%~dp0"
set "INV_PS=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"

if not exist "%~dp0Avvia Inventario.bat" goto :manca_avvio
if not exist "%~dp0Crea collegamento sul desktop.ps1" goto :manca_script
if not exist "%INV_PS%" goto :manca_powershell

"%INV_PS%" -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "%~dp0Crea collegamento sul desktop.ps1" %*
set "INV_ESITO=%errorlevel%"
if not "%INV_ESITO%"=="0" goto :errore
if not defined INV_SENZA_PAUSA pause
exit /b 0

:manca_avvio
echo Manca "Avvia Inventario.bat" in "%INV_DIR%".
goto :pacchetto_incompleto
:manca_script
echo Manca "Crea collegamento sul desktop.ps1" in "%INV_DIR%".
:pacchetto_incompleto
echo Estrai il pacchetto per intero nella stessa cartella e riprova.
set "INV_ESITO=1"
goto :errore
:manca_powershell
echo Windows PowerShell non e' disponibile su questa postazione.
set "INV_ESITO=1"
:errore
echo.
echo Il collegamento non e' stato creato. Leggi l'errore qui sopra.
echo Puoi crearlo a mano: tasto destro su "Avvia Inventario.bat",
echo Mostra altre opzioni, Invia a, Desktop ^(crea collegamento^).
if not defined INV_SENZA_PAUSA pause
exit /b %INV_ESITO%

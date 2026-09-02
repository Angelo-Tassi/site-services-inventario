@echo off
rem ---------------------------------------------------------------------------
rem  Crea sul desktop un collegamento al programma.
rem
rem  Il collegamento punta ad "Avvia Inventario.bat" e non direttamente a
rem  pythonw.exe: su alcune postazioni un collegamento a un eseguibile con
rem  argomenti viene rifiutato dai criteri di sicurezza, mentre quello a un file
rem  .bat funziona sempre. Parte minimizzato, cosi' la finestra nera si vede
rem  appena.
rem
rem  Il desktop viene chiesto a Windows, non costruito a mano: con OneDrive la
rem  scrivania vera e' dentro OneDrive, e %USERPROFILE%\Desktop e' una cartella
rem  che l'utente non vede mai.
rem ---------------------------------------------------------------------------
setlocal
set "INV_DIR=%~dp0"
set "INV_EXE=%~dp0Avvia Inventario.bat"
set "INV_ICON=%~dp0python\pythonw.exe,0"

if not exist "%INV_EXE%" (
    echo In questa cartella manca "Avvia Inventario.bat":
    echo   %INV_DIR%
    echo Estrai qui il pacchetto per intero e rilancia questo file.
    pause& exit /b 1
)
if not exist "%~dp0python\pythonw.exe" set "INV_ICON=%SystemRoot%\System32\shell32.dll,15"

powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$desk = [Environment]::GetFolderPath('Desktop');" ^
  "if (-not $desk -or -not (Test-Path $desk)) { $desk = Join-Path $env:USERPROFILE 'Desktop' };" ^
  "if (-not (Test-Path $desk)) { Write-Host 'Desktop non trovato.'; exit 1 };" ^
  "$lnk = Join-Path $desk 'Inventario dispositivi.lnk';" ^
  "$s = (New-Object -ComObject WScript.Shell).CreateShortcut($lnk);" ^
  "$s.TargetPath = $env:INV_EXE; $s.WorkingDirectory = $env:INV_DIR;" ^
  "$s.IconLocation = $env:INV_ICON; $s.WindowStyle = 7;" ^
  "$s.Description = 'Inventario laptop e tablet'; $s.Save();" ^
  "Copy-Item $lnk (Join-Path $env:INV_DIR 'Inventario dispositivi.lnk') -Force;" ^
  "Write-Host ''; Write-Host 'Collegamento creato qui:'; Write-Host ('  ' + $lnk)"

echo.
echo --------------------------------------------------------------------
echo  SE SUL DESKTOP NON COMPARE NIENTE
echo.
echo  Su alcune postazioni la creazione automatica viene bloccata. In quel
echo  caso il collegamento si fa a mano, e funziona sempre:
echo.
echo     tasto destro su  "Avvia Inventario.bat"  in questa cartella
echo     ^> Mostra altre opzioni ^> Invia a ^> Desktop ^(crea collegamento^)
echo.
echo  In questa cartella e' rimasta anche una copia gia' pronta,
echo  "Inventario dispositivi.lnk": si puo' trascinare sul desktop, anche
echo  su quello di un altro utente, senza eseguire niente.
echo --------------------------------------------------------------------
echo.
pause

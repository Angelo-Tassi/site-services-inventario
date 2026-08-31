@echo off
rem ---------------------------------------------------------------------------
rem  Crea sul desktop un collegamento al programma che sta in questa cartella.
rem  Il desktop viene chiesto a Windows, non costruito a mano: con OneDrive la
rem  scrivania vera e' dentro OneDrive, e %USERPROFILE%\Desktop e' una cartella
rem  che l'utente non vede mai.
rem ---------------------------------------------------------------------------
setlocal
set "INV_DIR=%~dp0"

if not exist "%~dp0python\pythonw.exe" (
    echo In questa cartella manca la sottocartella python:
    echo   %INV_DIR%
    echo Estrai qui il pacchetto per intero e rilancia questo file.
    pause& exit /b 1
)
set "INV_EXE=%~dp0python\pythonw.exe"
set "INV_ARGS=""%~dp0Inventario.py"""
set "INV_ICON=%~dp0python\pythonw.exe,0"

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
    echo Puoi farlo a mano: tasto destro su "Avvia Inventario.bat"
    echo ^> Mostra altre opzioni ^> Invia a ^> Desktop ^(crea collegamento^).
    pause& exit /b 1
)

echo Una copia e' rimasta nella cartella del programma: si puo' trascinare sul
echo desktop di un altro utente senza eseguire nulla.
echo.
pause

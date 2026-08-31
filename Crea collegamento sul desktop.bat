@echo off
rem ---------------------------------------------------------------------------
rem  Crea sul desktop un collegamento al programma che sta in questa cartella.
rem  Funziona con tutti e due i pacchetti:
rem    - quello con Inventario.exe, che viene avviato direttamente;
rem    - quello senza eseguibile nostro, dove ad avviare il programma e' il
rem      pythonw.exe ufficiale che sta nella sottocartella "python".
rem  Il collegamento viene salvato anche qui accanto, cosi' puo' essere copiato
rem  sul desktop degli altri utenti senza eseguire nulla.
rem ---------------------------------------------------------------------------
setlocal
set "INV_DIR=%~dp0"
set "INV_LNK=%USERPROFILE%\Desktop\Inventario dispositivi.lnk"
set "INV_ARGS="

if exist "%~dp0Inventario.exe" (
    set "INV_EXE=%~dp0Inventario.exe"
    set "INV_ICON=%~dp0Inventario.exe,0"
) else if exist "%~dp0python\pythonw.exe" (
    set "INV_EXE=%~dp0python\pythonw.exe"
    set "INV_ARGS=""%~dp0Inventario.py"""
    set "INV_ICON=%~dp0python\pythonw.exe,0"
) else (
    echo In questa cartella non c'e' ne' Inventario.exe ne' la cartella python:
    echo   %INV_DIR%
    echo Estrai qui il pacchetto per intero e rilancia questo file.
    pause& exit /b 1
)

powershell -NoProfile -ExecutionPolicy Bypass -Command "$s = (New-Object -ComObject WScript.Shell).CreateShortcut($env:INV_LNK); $s.TargetPath = $env:INV_EXE; $s.Arguments = $env:INV_ARGS; $s.WorkingDirectory = $env:INV_DIR; $s.IconLocation = $env:INV_ICON; $s.Description = 'Inventario laptop e tablet'; $s.Save()"
if errorlevel 1 (echo Creazione del collegamento non riuscita.& pause& exit /b 1)

copy /y "%INV_LNK%" "%~dp0Inventario dispositivi.lnk" >nul 2>&1

echo.
echo Collegamento creato sul desktop: "Inventario dispositivi"
echo Una copia e' rimasta nella cartella di rete: gli altri utenti possono
echo trascinarla sul proprio desktop, senza eseguire nulla.
echo.
pause

@echo off
rem ---------------------------------------------------------------------------
rem  Crea sul desktop un collegamento a Inventario.exe che sta in questa
rem  cartella. Il collegamento viene salvato anche qui accanto, cosi' puo'
rem  essere copiato sul desktop degli altri utenti senza eseguire nulla.
rem ---------------------------------------------------------------------------
setlocal
set "INV_DIR=%~dp0"
set "INV_EXE=%~dp0Inventario.exe"
set "INV_LNK=%USERPROFILE%\Desktop\Inventario dispositivi.lnk"

if not exist "%INV_EXE%" (
    echo Inventario.exe non e' in questa cartella:
    echo   %INV_DIR%
    echo Copia qui l'eseguibile e rilancia questo file.
    pause& exit /b 1
)

powershell -NoProfile -ExecutionPolicy Bypass -Command "$s = (New-Object -ComObject WScript.Shell).CreateShortcut($env:INV_LNK); $s.TargetPath = $env:INV_EXE; $s.WorkingDirectory = $env:INV_DIR; $s.IconLocation = $env:INV_EXE + ',0'; $s.Description = 'Inventario laptop e tablet'; $s.Save()"
if errorlevel 1 (echo Creazione del collegamento non riuscita.& pause& exit /b 1)

copy /y "%INV_LNK%" "%~dp0Inventario dispositivi.lnk" >nul 2>&1

echo.
echo Collegamento creato sul desktop: "Inventario dispositivi"
echo Una copia e' rimasta nella cartella di rete: gli altri utenti possono
echo trascinarla sul proprio desktop, senza eseguire nulla.
echo.
pause

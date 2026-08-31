@echo off
rem ---------------------------------------------------------------------------
rem  Raccoglie le informazioni sull'installazione e le scrive in
rem  Diagnostica.txt, qui accanto. Non modifica niente.
rem  Doppio clic, poi manda il file a chi ti assiste.
rem ---------------------------------------------------------------------------
setlocal
pushd "%~dp0" || (echo Impossibile raggiungere la cartella del programma. & pause & exit /b 1)

echo.
echo Raccolta delle informazioni in corso...
echo.

if exist "%~dp0python\python.exe" (
    "%~dp0python\python.exe" -m inventario.diagnostica
) else (
    py -3 -m inventario.diagnostica 2>nul || python -m inventario.diagnostica
)
if errorlevel 1 (
    echo.
    echo Non e' stato possibile eseguire la diagnostica.
    echo Se in questa cartella c'e' solo Inventario.exe, usa il pacchetto
    echo senza eseguibile: la diagnostica ha bisogno dei file .py.
)

rem  la finestra del blocco note aiuta a ritrovare il file da allegare
if exist "%~dp0Diagnostica.txt" start "" notepad "%~dp0Diagnostica.txt"

echo.
pause
popd

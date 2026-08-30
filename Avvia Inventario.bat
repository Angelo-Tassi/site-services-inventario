@echo off
rem Avvio dai sorgenti (richiede Python + openpyxl sul PC).
rem Per la cartella di rete usare invece Inventario.exe: non richiede nulla.
pushd "%~dp0" || (echo Impossibile raggiungere la cartella del programma. & pause & exit /b 1)
where pythonw >nul 2>&1
if %errorlevel%==0 (
    start "" pythonw "Inventario.py"
) else (
    py -w "Inventario.py" 2>nul || python "Inventario.py"
)
popd

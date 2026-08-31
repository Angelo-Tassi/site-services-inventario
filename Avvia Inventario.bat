@echo off
rem ---------------------------------------------------------------------------
rem  Avvia il programma. Se accanto c'e' la cartella "python" usa quello, che
rem  viaggia col programma e non richiede niente sul PC. Altrimenti ripiega sul
rem  Python installato sulla macchina.
rem ---------------------------------------------------------------------------
pushd "%~dp0" || (echo Impossibile raggiungere la cartella del programma. & pause & exit /b 1)
if exist "%~dp0python\pythonw.exe" (
    start "" "%~dp0python\pythonw.exe" "%~dp0Inventario.py"
    popd
    exit /b 0
)
where pythonw >nul 2>&1
if %errorlevel%==0 (
    start "" pythonw "Inventario.py"
) else (
    py -w "Inventario.py" 2>nul || python "Inventario.py"
)
popd

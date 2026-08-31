@echo off
rem ---------------------------------------------------------------------------
rem  Dice al programma quale inventario condiviso deve aprire.
rem  Si esegue una volta sola: quello che scrive viaggia con la cartella, che
rem  si puo' poi copiare gia' configurata su tutte le postazioni.
rem ---------------------------------------------------------------------------
setlocal
pushd "%~dp0" || (echo Impossibile raggiungere la cartella del programma. & pause & exit /b 1)
if exist "%~dp0python\python.exe" (
    "%~dp0python\python.exe" -m inventario.configura
) else (
    py -3 -m inventario.configura 2>nul || python -m inventario.configura
)
echo.
pause
popd

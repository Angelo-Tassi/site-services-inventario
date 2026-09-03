@echo off
rem ---------------------------------------------------------------------------
rem  Dice al programma quale inventario condiviso deve aprire.
rem  Si esegue una volta sola: quello che scrive viaggia con la cartella, che
rem  si puo' poi copiare gia' configurata su tutte le postazioni.
rem ---------------------------------------------------------------------------
rem  Il primo doppio clic su un file di questo pacchetto, appena estratto da
rem  uno zip scaricato, fa comparire l'avviso "Editore sconosciuto": e' il
rem  contrassegno che Windows mette su tutto cio' che arriva da Internet, non
rem  ha a che fare con una firma mancante. Qui lo si toglie da tutta la
rem  cartella una volta sola, cosi' i prossimi doppi clic - su questo file e
rem  sugli altri accanto - non lo chiedono piu'.
if exist "%~dp0.sbloccato" goto :gia_sbloccato
powershell -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -Command "Get-ChildItem -LiteralPath '%~dp0' -Recurse -File -ErrorAction SilentlyContinue | Unblock-File -ErrorAction SilentlyContinue" >nul 2>&1
type nul > "%~dp0.sbloccato" 2>nul
attrib +h "%~dp0.sbloccato" >nul 2>&1
:gia_sbloccato
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

@echo off
rem ---------------------------------------------------------------------------
rem  Avvia il programma. Se accanto c'e' la cartella "python" usa quello, che
rem  viaggia col programma e non richiede niente sul PC. Altrimenti ripiega sul
rem  Python installato sulla macchina.
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

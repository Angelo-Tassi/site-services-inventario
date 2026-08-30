@echo off
rem ---------------------------------------------------------------------------
rem  Da eseguire UNA VOLTA SOLA, su un qualsiasi PC Windows con Python
rem  installato. Produce Distribuzione\Inventario.exe: un singolo file che
rem  contiene Python, le librerie e il programma. Va copiato nella cartella di
rem  rete e non richiede nulla sui PC che lo useranno.
rem ---------------------------------------------------------------------------
setlocal
pushd "%~dp0"
if errorlevel 1 (echo Impossibile raggiungere la cartella del programma.& pause& exit /b 1)

where python >nul 2>&1
if errorlevel 1 (
    echo Python non trovato.
    echo Installalo da https://www.python.org/downloads/windows/ ricordando di
    echo spuntare "Add Python to PATH", poi rilancia questo file.
    pause& popd& exit /b 1
)

echo.
echo [1/3] Installazione degli strumenti di compilazione...
python -m pip install --upgrade --quiet pyinstaller openpyxl
if errorlevel 1 (echo Installazione non riuscita.& pause& popd& exit /b 1)

echo [2/3] Compilazione dell'eseguibile (puo' richiedere qualche minuto)...
python -m PyInstaller --noconfirm --clean --onefile --noconsole --name Inventario --collect-submodules openpyxl Inventario.py
if errorlevel 1 (echo Compilazione non riuscita.& pause& popd& exit /b 1)

echo [3/3] Preparazione della cartella da copiare in rete...
if not exist "Distribuzione" mkdir "Distribuzione"
copy /y "dist\Inventario.exe" "Distribuzione\Inventario.exe" >nul
copy /y "Crea collegamento sul desktop.bat" "Distribuzione\" >nul
copy /y "README.md" "Distribuzione\Come funziona.txt" >nul
rem i file di collaudo viaggiano con il programma: servono a provarlo in sede
if not exist "Distribuzione\Collaudo" mkdir "Distribuzione\Collaudo"
copy /y "Collaudo\*.xlsx" "Distribuzione\Collaudo\" >nul
copy /y "Collaudo\README.md" "Distribuzione\Collaudo\" >nul
copy /y "Collaudo\README.en.md" "Distribuzione\Collaudo\" >nul

echo.
echo ===========================================================================
echo  Fatto. Copia il CONTENUTO della cartella "Distribuzione" nella cartella
echo  di rete condivisa. Al primo avvio il programma crea li' Inventario.xlsx.
echo ===========================================================================
echo.
popd
pause

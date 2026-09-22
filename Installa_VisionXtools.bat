@echo off
setlocal
title Installazione VisionXtools
set "EXTDIR=%APPDATA%\pyRevit\Extensions\VisionXtools.extension"
set "REPO=https://github.com/VisionXt-tech/VisionXtools.extension.git"

echo.
echo ====================================================
echo   Installazione VisionXtools per pyRevit
echo ====================================================
echo.

where pyrevit >nul 2>&1
if errorlevel 1 (
    echo ERRORE: pyRevit non risulta installato su questo PC.
    echo Installa prima pyRevit da https://pyrevitlabs.io e poi rilancia questo file.
    echo.
    pause
    exit /b 1
)

if not exist "%EXTDIR%" goto :installa

echo VisionXtools e' gia' presente in:
echo    %EXTDIR%
echo.
echo Questa procedura la sostituisce con l'ultima versione da GitHub.
echo La copia attuale viene rinominata, non cancellata.
echo.
echo IMPORTANTE: chiudi Revit prima di procedere.
echo.
choice /c SN /n /m "Procedere? [S = si, N = no] "
if errorlevel 2 (
    echo.
    echo Operazione annullata, non e' stato modificato nulla.
    echo.
    pause
    exit /b 0
)

echo.
set "BACKUP=%EXTDIR%.old_%RANDOM%"
move "%EXTDIR%" "%BACKUP%" >nul 2>&1
if errorlevel 1 (
    echo ERRORE: non riesco a spostare la cartella, e' in uso.
    echo Chiudi Revit e rilancia questo file.
    echo.
    pause
    exit /b 1
)
echo Copia precedente salvata in:
echo    %BACKUP%

:installa
echo.
echo Scarico l'estensione, puo' richiedere un minuto...
echo Eventuali messaggi di errore sul file di configurazione sono normali.
echo.
pyrevit extend ui VisionXtools "%REPO%" --branch=main >nul 2>&1

echo.
if exist "%EXTDIR%\VisionXtools.tab" goto :riuscita

REM ponytail: on failure just put the old copy back, no partial repair
if not defined BACKUP goto :fallita
if not exist "%BACKUP%" goto :fallita
rd /s /q "%EXTDIR%" >nul 2>&1
move "%BACKUP%" "%EXTDIR%" >nul 2>&1
echo Ripristinata la copia precedente.

:fallita
echo ====================================================
echo   INSTALLAZIONE NON RIUSCITA
echo ====================================================
echo.
echo Riprova, oppure segui la procedura manuale qui:
echo https://github.com/VisionXt-tech/VisionXtools.extension
echo.
pause
exit /b 1

:riuscita
echo ====================================================
echo   INSTALLAZIONE COMPLETATA
echo ====================================================
echo.
echo Apri Revit: troverai il tab VisionXtools aggiornato.
echo Se Revit e' gia' aperto: tab pyRevit ^> Reload.
echo.
if defined BACKUP echo Se e' tutto a posto puoi cancellare la cartella .old_ di backup.
echo.
pause
exit /b 0

@echo off
setlocal
title Installazione VisionXtools
set "EXTDIR=%APPDATA%\pyRevit\Extensions\VisionXtools.extension"

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

if exist "%EXTDIR%\VisionXtools.tab" (
    echo VisionXtools e' gia' installata in:
    echo    %EXTDIR%
    echo.
    echo Per aggiornarla, in Revit: tab VisionXtools ^> About ^> Update.
    echo.
    pause
    exit /b 0
)

echo Scarico l'estensione, puo' richiedere un minuto...
echo Eventuali messaggi di errore sul file di configurazione sono normali.
echo.
pyrevit extend ui VisionXtools https://github.com/VisionXt-tech/VisionXtools.extension.git --branch=main >nul 2>&1

echo.
if exist "%EXTDIR%\VisionXtools.tab" (
    echo ====================================================
    echo   INSTALLAZIONE COMPLETATA
    echo ====================================================
    echo.
    echo Apri Revit: troverai il tab VisionXtools.
    echo Se Revit e' gia' aperto: tab pyRevit ^> Reload.
    echo.
    pause
    exit /b 0
) else (
    echo ====================================================
    echo   INSTALLAZIONE NON RIUSCITA
    echo ====================================================
    echo.
    echo Riprova, oppure segui la procedura manuale qui:
    echo https://github.com/VisionXt-tech/VisionXtools.extension
    echo.
    pause
    exit /b 1
)

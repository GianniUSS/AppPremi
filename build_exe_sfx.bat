@echo off
setlocal enableextensions enabledelayedexpansion

set SCRIPT_DIR=%~dp0
pushd "%SCRIPT_DIR%"

set APP_NAME=GestionePremi
set SOURCE_DIR=%SCRIPT_DIR%dist\%APP_NAME%
set SFX_WORKDIR=%SCRIPT_DIR%dist\_sfx
set SFX_TARGET=%SCRIPT_DIR%dist\%APP_NAME%.exe
set SEVEN_ZIP=C:\Program Files\7-Zip\7z.exe
set SEVEN_SFX=C:\Program Files\7-Zip\7z.sfx
set SFX_CONFIG=%SFX_WORKDIR%\sfx_config.txt
set SFX_ARCHIVE=%SFX_WORKDIR%\payload.7z

REM Build con PyInstaller (directory mode)
call "%SCRIPT_DIR%build_exe.bat"
if %ERRORLEVEL% neq 0 goto :end

if exist "%SFX_TARGET%" del /f "%SFX_TARGET%"
if exist "%SFX_WORKDIR%" rmdir /s /q "%SFX_WORKDIR%"
mkdir "%SFX_WORKDIR%"

if not exist "%SEVEN_ZIP%" (
    echo [ERRORE] 7-Zip non trovato. Installa 7-Zip per creare l'SFX.
    goto :end
)
if not exist "%SEVEN_SFX%" (
    echo [ERRORE] Modulo 7z.sfx non trovato.
    goto :end
)

REM Crea script di installazione che viene incluso nell'archivio
(
    echo @echo off
    echo setlocal
    echo set "TARGET=%%LOCALAPPDATA%%\AppEden\GestionePremi"
    echo.
    echo REM Chiudi eventuale istanza in esecuzione
    echo taskkill /F /IM GestionePremi.exe 2^>nul
    echo timeout /t 2 /nobreak ^>nul
    echo.
    echo REM Rimuovi vecchia installazione e copia i nuovi file
    echo if exist "%%TARGET%%" rmdir /s /q "%%TARGET%%"
    echo mkdir "%%TARGET%%"
    echo xcopy /E /I /Y "%%~dp0*" "%%TARGET%%\" ^>nul
    echo del "%%TARGET%%\install.cmd" 2^>nul
    echo.
    echo REM Avvia la nuova versione
    echo start "" "%%TARGET%%\GestionePremi.exe"
    echo endlocal
) > "%SOURCE_DIR%\install.cmd"

REM Crea archivio 7z con tutto il contenuto + install.cmd
if exist "%SFX_ARCHIVE%" del /f "%SFX_ARCHIVE%"
"%SEVEN_ZIP%" a -t7z -mx=5 "%SFX_ARCHIVE%" "%SOURCE_DIR%\*" >nul

REM Configurazione SFX - estrae in temp ed esegue install.cmd
(
    echo ;!@Install@!UTF-8!
    echo Title="Installazione %APP_NAME%"
    echo RunProgram="install.cmd"
    echo ;!@InstallEnd@!
) > "%SFX_CONFIG%"

REM Crea l'eseguibile SFX finale
copy /b "%SEVEN_SFX%" + "%SFX_CONFIG%" + "%SFX_ARCHIVE%" "%SFX_TARGET%" >nul

REM Pulisci
del "%SOURCE_DIR%\install.cmd" 2>nul

if not exist "%SFX_TARGET%" (
    echo [ERRORE] Creazione SFX fallita.
    goto :end
)

echo SFX creato: %SFX_TARGET%

:end
popd
endlocal
exit /b %ERRORLEVEL%

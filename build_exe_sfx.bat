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

if exist "%SFX_ARCHIVE%" del /f "%SFX_ARCHIVE%"
"%SEVEN_ZIP%" a -t7z "%SFX_ARCHIVE%" "%SOURCE_DIR%\*" >nul

(
    echo @echo off
    echo setlocal
    echo set "TARGET=%%LOCALAPPDATA%%\AppEden\GestionePremi"
    echo if exist "%%TARGET%%" rmdir /s /q "%%TARGET%%"
    echo mkdir "%%TARGET%%"
    echo powershell -NoProfile -ExecutionPolicy Bypass -Command "Expand-Archive -Force -Path '%%~dp0payload.7z' -DestinationPath '%%TARGET%%'"
    echo start "" "%%TARGET%%\GestionePremi.exe"
    echo endlocal
) > "%SFX_WORKDIR%\run_update.cmd"

(
    echo ;!@Install@!UTF-8!
    echo Title="%APP_NAME%"
    echo RunProgram="run_update.cmd"
    echo ;!@InstallEnd@!
) > "%SFX_CONFIG%"

copy /b "%SEVEN_SFX%" + "%SFX_CONFIG%" + "%SFX_ARCHIVE%" "%SFX_TARGET%" >nul

if not exist "%SFX_TARGET%" (
    echo [ERRORE] Creazione SFX fallita.
    goto :end
)

echo SFX creato: %SFX_TARGET%

:end
popd
endlocal
exit /b %ERRORLEVEL%

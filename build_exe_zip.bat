@echo off
setlocal enableextensions enabledelayedexpansion

set SCRIPT_DIR=%~dp0
pushd "%SCRIPT_DIR%"

set APP_NAME=GestionePremi
set SOURCE_DIR=%SCRIPT_DIR%dist\%APP_NAME%
set ZIP_TARGET=%SCRIPT_DIR%dist\%APP_NAME%.zip
set SEVEN_ZIP=C:\Program Files\7-Zip\7z.exe

REM Build con PyInstaller (directory mode)
call "%SCRIPT_DIR%build_exe.bat"
if %ERRORLEVEL% neq 0 goto :end

if exist "%ZIP_TARGET%" del /f "%ZIP_TARGET%"

if not exist "%SEVEN_ZIP%" (
    echo [ERRORE] 7-Zip non trovato. Installa 7-Zip.
    goto :end
)

REM Crea ZIP con la struttura GestionePremi\* dentro
cd dist
"%SEVEN_ZIP%" a -tzip "%ZIP_TARGET%" "%APP_NAME%" >nul
cd ..

if not exist "%ZIP_TARGET%" (
    echo [ERRORE] Creazione ZIP fallita.
    goto :end
)

echo ZIP creato: %ZIP_TARGET%

:end
popd
endlocal
exit /b %ERRORLEVEL%

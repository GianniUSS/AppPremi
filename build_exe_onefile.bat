@echo off
setlocal enableextensions enabledelayedexpansion

set SCRIPT_DIR=%~dp0
pushd "%SCRIPT_DIR%"

set PYTHON_BIN=.venv\Scripts\python.exe
set PYINSTALLER_MODULE=PyInstaller
set MAIN_SCRIPT=main.py
set APP_NAME=GestionePremi
set APP_NAME_ONEFILE=%APP_NAME%_onefile
set MYSQL_PLUGIN_SRC=.venv\Lib\site-packages\mysql\vendor\plugin\*.dll
set MYSQL_PLUGIN_DST=mysql/vendor/plugin

if not exist "%PYTHON_BIN%" (
    echo [ERRORE] Interpreter Python non trovato in %%PYTHON_BIN%%
    echo          Attiva il virtualenv o reinstalla le dipendenze.
    goto :end
)

echo ==============================================
echo   Generazione eseguibile ONEFILE %APP_NAME%
echo   Cartella progetto: %SCRIPT_DIR%
echo ==============================================

if exist dist\%APP_NAME_ONEFILE%.exe (
    echo [INFO] Rimozione dist\%APP_NAME_ONEFILE%.exe...
    del /f dist\%APP_NAME_ONEFILE%.exe
)
if exist dist\%APP_NAME%.exe (
    echo [INFO] Rimozione dist\%APP_NAME%.exe...
    del /f dist\%APP_NAME%.exe
)
if exist build\%APP_NAME_ONEFILE% (
    echo [INFO] Pulizia directory build\%APP_NAME_ONEFILE%...
    rmdir /s /q build\%APP_NAME_ONEFILE%
)
if exist %APP_NAME_ONEFILE%.spec (
    echo [INFO] Rimozione file %APP_NAME_ONEFILE%.spec...
    del /f %APP_NAME_ONEFILE%.spec
)

echo [INFO] Avvio PyInstaller ONEFILE (tramite python -m PyInstaller)...
call "%PYTHON_BIN%" -m %PYINSTALLER_MODULE% --onefile --noconsole --clean --name %APP_NAME_ONEFILE% "%MAIN_SCRIPT%" ^
    --collect-submodules mysql.connector.plugins ^
    --add-binary "%MYSQL_PLUGIN_SRC%";%MYSQL_PLUGIN_DST%

set EXIT_CODE=%ERRORLEVEL%
if %EXIT_CODE% neq 0 (
    echo [ERRORE] Build ONEFILE fallita con codice %EXIT_CODE%
    goto :end
)

echo Build ONEFILE completata con successo. Output: dist\%APP_NAME_ONEFILE%.exe
if exist dist\%APP_NAME_ONEFILE%.exe (
    echo [INFO] Copia/Rinomina per release: dist\%APP_NAME%.exe
    copy /y dist\%APP_NAME_ONEFILE%.exe dist\%APP_NAME%.exe >nul
)
echo ==============================================

:end
popd
endlocal
exit /b %EXIT_CODE%

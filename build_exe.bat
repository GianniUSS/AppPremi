@echo off
setlocal enableextensions enabledelayedexpansion

set SCRIPT_DIR=%~dp0
pushd "%SCRIPT_DIR%"

set PYTHON_BIN=.venv\Scripts\python.exe
set PYINSTALLER_MODULE=PyInstaller
set MAIN_SCRIPT=main.py
set APP_NAME=GestionePremi
set MYSQL_PLUGIN_SRC=.venv\Lib\site-packages\mysql\vendor\plugin\*.dll
set MYSQL_PLUGIN_DST=mysql/vendor/plugin

if not exist "%PYTHON_BIN%" (
    echo [ERRORE] Interpreter Python non trovato in %%PYTHON_BIN%%
    echo          Attiva il virtualenv o reinstalla le dipendenze.
    goto :end
)

echo ==============================================
echo   Generazione eseguibile %APP_NAME%
echo   Cartella progetto: %SCRIPT_DIR%
echo ==============================================

if exist dist\%APP_NAME% (
    echo [INFO] Pulizia directory dist\%APP_NAME%...
    rmdir /s /q dist\%APP_NAME%
)
if exist build\%APP_NAME% (
    echo [INFO] Pulizia directory build\%APP_NAME%...
    rmdir /s /q build\%APP_NAME%
)
if exist %APP_NAME%.spec (
    echo [INFO] Rimozione file %APP_NAME%.spec%...
    del /f %APP_NAME%.spec
)

echo [INFO] Avvio PyInstaller (tramite python -m PyInstaller)...
call "%PYTHON_BIN%" -m %PYINSTALLER_MODULE% --noconsole --clean --name %APP_NAME% "%MAIN_SCRIPT%" ^
    --collect-submodules mysql.connector.plugins ^
    --add-binary "%MYSQL_PLUGIN_SRC%";%MYSQL_PLUGIN_DST%

set EXIT_CODE=%ERRORLEVEL%
if %EXIT_CODE% neq 0 (
    echo [ERRORE] Build fallita con codice %EXIT_CODE%
    goto :end
)

echo Build completata con successo. Output in dist\%APP_NAME%
echo ==============================================

:end
popd
endlocal
exit /b %EXIT_CODE%

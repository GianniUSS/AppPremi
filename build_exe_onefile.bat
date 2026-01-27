@echo off
setlocal enableextensions enabledelayedexpansion

set SCRIPT_DIR=%~dp0
pushd "%SCRIPT_DIR%"

set PYTHON_BIN=.venv\Scripts\python.exe
set PYINSTALLER_MODULE=PyInstaller
set MAIN_SCRIPT=main.py
set APP_NAME=GestionePremi
set APP_NAME_ONEFILE=%APP_NAME%
set MYSQL_PLUGIN_SRC=.venv\Lib\site-packages\mysql\vendor\plugin\*.dll
set MYSQL_PLUGIN_DST=mysql/vendor/plugin
set PYTHON_DLL=
set ADD_PYDLL=
set PYTHON_DLL_TMP=%TEMP%\pythondll.txt
set VCRUNTIME_DLL_TMP=%TEMP%\vcruntimedlls.txt
set ADD_VCRUNTIME=

if exist "%PYTHON_DLL_TMP%" del /f "%PYTHON_DLL_TMP%"
"%PYTHON_BIN%" -c "import sys, os; print(os.path.join(sys.base_prefix, 'python{}{}.dll'.format(sys.version_info[0], sys.version_info[1])))" > "%PYTHON_DLL_TMP%"
set /p PYTHON_DLL=<"%PYTHON_DLL_TMP%"
if exist "%PYTHON_DLL%" (
    set ADD_PYDLL=--add-binary "%PYTHON_DLL%";.
)

if exist "%VCRUNTIME_DLL_TMP%" del /f "%VCRUNTIME_DLL_TMP%"
"%PYTHON_BIN%" -c "import os, sys; names=('vcruntime140.dll','vcruntime140_1.dll'); paths=[sys.base_prefix, os.path.join(sys.base_prefix,'DLLs'), os.path.join(os.environ.get('SystemRoot','C:\\Windows'),'System32')]; found=[next((os.path.join(p,n) for p in paths if os.path.exists(os.path.join(p,n))), '') for n in names]; print('\n'.join([p for p in found if p]))" > "%VCRUNTIME_DLL_TMP%"
for /f "usebackq delims=" %%A in ("%VCRUNTIME_DLL_TMP%") do (
    if exist "%%A" set "ADD_VCRUNTIME=!ADD_VCRUNTIME! --add-binary ^"%%A^";."
)

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
    --icon "assets\app_icon.ico" ^
    --collect-submodules mysql.connector.plugins ^
    --add-binary "%MYSQL_PLUGIN_SRC%";%MYSQL_PLUGIN_DST% ^
    %ADD_PYDLL% %ADD_VCRUNTIME%

set EXIT_CODE=%ERRORLEVEL%
if %EXIT_CODE% neq 0 (
    echo [ERRORE] Build ONEFILE fallita con codice %EXIT_CODE%
    goto :end
)

echo Build ONEFILE completata con successo. Output: dist\%APP_NAME_ONEFILE%.exe
echo ==============================================

:end
popd
endlocal
exit /b %EXIT_CODE%

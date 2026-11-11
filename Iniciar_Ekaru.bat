@echo off
setlocal
cd /d "%~dp0"
set "PYTHONW="
if exist "C:\Python314\pythonw.exe" set "PYTHONW=C:\Python314\pythonw.exe"
if not defined PYTHONW for /f "delims=" %%I in ('where pythonw.exe 2^>nul') do if not defined PYTHONW set "PYTHONW=%%I"
if not defined PYTHONW goto :nopython
set "PYTHONC=%PYTHONW:pythonw.exe=python.exe%"
"%PYTHONC%" -c "import customtkinter" >nul 2>&1
if errorlevel 1 goto :install
:run
start "" /D "%~dp0" "%PYTHONW%" "%~dp0app.py"
exit /b 0
:install
echo Falta el modulo customtkinter. Se instala una sola vez (requiere internet)...
"%PYTHONC%" -m pip install customtkinter
"%PYTHONC%" -c "import customtkinter" >nul 2>&1
if errorlevel 1 goto :fail
goto :run
:nopython
echo No se encontro Python (pythonw.exe). Instala Python 3 desde python.org y marca "Add to PATH".
pause
exit /b 1
:fail
echo No se pudo instalar customtkinter. Revisa la conexion a internet.
pause
exit /b 1

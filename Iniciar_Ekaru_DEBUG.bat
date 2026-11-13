@echo off
setlocal
cd /d "%~dp0"
set "PYTHONC="
if exist "C:\Python314\python.exe" set "PYTHONC=C:\Python314\python.exe"
if not defined PYTHONC for /f "delims=" %%I in ('where python.exe 2^>nul') do if not defined PYTHONC set "PYTHONC=%%I"
if not defined PYTHONC (
  echo No se encontro Python.
  pause
  exit /b 1
)
echo Usando: %PYTHONC%
"%PYTHONC%" --version
"%PYTHONC%" -c "import customtkinter; print('customtkinter', customtkinter.__version__)"
echo.
echo Iniciando Etiko con consola visible. Los errores aparecen aqui.
"%PYTHONC%" "%~dp0app.py"
echo.
echo La aplicacion se cerro (codigo %errorlevel%).
pause

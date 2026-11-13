@echo off
setlocal
if not exist "%LOCALAPPDATA%\Etiko" mkdir "%LOCALAPPDATA%\Etiko"
copy /y "%~dp0etiko.ico" "%LOCALAPPDATA%\Etiko\etiko.ico" >nul
powershell -NoProfile -ExecutionPolicy Bypass -Command "$d=[Environment]::GetFolderPath('DesktopDirectory'); $s=(New-Object -ComObject WScript.Shell).CreateShortcut($d+'\Etiko.lnk'); $s.TargetPath='%~dp0Iniciar_Ekaru.bat'; $s.WorkingDirectory='%~dp0'; $s.IconLocation=$env:LOCALAPPDATA+'\Etiko\etiko.ico'; $s.WindowStyle=7; $s.Description='Etiko - etiquetas Glace'; $s.Save(); Write-Host ('Icono creado en: ' + $d)"
echo.
pause

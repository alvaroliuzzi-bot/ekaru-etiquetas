@echo off
setlocal
if not exist "%LOCALAPPDATA%\EKaru" mkdir "%LOCALAPPDATA%\EKaru"
copy /y "%~dp0ekaru.ico" "%LOCALAPPDATA%\EKaru\ekaru.ico" >nul
powershell -NoProfile -ExecutionPolicy Bypass -Command "$d=[Environment]::GetFolderPath('DesktopDirectory'); $s=(New-Object -ComObject WScript.Shell).CreateShortcut($d+'\E-Karu.lnk'); $s.TargetPath='%~dp0Iniciar_Ekaru.bat'; $s.WorkingDirectory='%~dp0'; $s.IconLocation=$env:LOCALAPPDATA+'\EKaru\ekaru.ico'; $s.WindowStyle=7; $s.Description='E-Karu Etiquetas'; $s.Save(); Write-Host ('Icono creado en: ' + $d)"
echo.
pause

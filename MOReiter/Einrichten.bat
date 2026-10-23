@echo off
rem Richtet den MOReiter dieses Ordners auf diesem PC ein: beendet den laufenden MOReiter (Tray) und startet die
rem exe aus diesem Ordner neu, z.B. nachdem sie auf einem anderen PC kompiliert wurde.
rem   Einrichten.bat           beenden und neu starten
rem   Einrichten.bat beenden   nur beenden (Kompilieren.bat, die exe laesst sich nicht ueberschreiben, solange sie laeuft)
rem   Einrichten.bat starten   nur starten
rem Achtung: beendet auch eine laufende Smart Pix-Liste (Strg+Umsch+S)
cd /d "%~dp0"
if /i "%~1"=="starten" goto starten
tasklist /fi "imagename eq MOReiter.exe" 2>nul | find /i "MOReiter.exe" >nul
if errorlevel 1 goto beendet
echo Beende den laufenden MOReiter ...
taskkill /im MOReiter.exe >nul 2>&1
for /l %%i in (1,1,10) do (
	tasklist /fi "imagename eq MOReiter.exe" 2>nul | find /i "MOReiter.exe" >nul || goto beendet
	timeout /t 1 /nobreak >nul
)
taskkill /f /im MOReiter.exe >nul 2>&1
timeout /t 1 /nobreak >nul
:beendet
if /i "%~1"=="beenden" goto :eof
:starten
if exist MOReiter.exe start "" "%~dp0MOReiter.exe"

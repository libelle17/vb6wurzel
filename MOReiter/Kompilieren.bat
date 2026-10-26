@echo off
rem Kompiliert MOReiter.au3 mit dem Reiter-Icon; beendet vorher den laufenden MOReiter dieses PCs (Tray) und
rem startet danach den neuen (Einrichten.bat; auf anderen PCs genuegt danach Einrichten.bat allein)
rem Achtung: beendet auch eine laufende Smart Pix-Liste (Strg+Umsch+S)
rem Gibt es noch keine MOReiter.exe (erstes Kompilieren), fragt "MOReiter.exe Praxis" danach, was die Praxis nutzt
rem (Labor, Zuckerportale, Smart Pix, Patientenlaufzettel), und passt MOReiter-Vorgaben.ini daran an
cd /d "%~dp0"
set A2E=%ProgramFiles(x86)%\AutoIt3\Aut2Exe\Aut2exe.exe
if not exist "%A2E%" set A2E=%ProgramFiles%\AutoIt3\Aut2Exe\Aut2exe.exe
set ERST=
if not exist MOReiter.exe set ERST=1
call Einrichten.bat beenden
"%A2E%" /in MOReiter.au3 /out MOReiter.exe /icon MOReiter.ico /x86
if errorlevel 1 (echo Kompilieren fehlgeschlagen & pause)
if defined ERST if exist MOReiter.exe start "" /wait MOReiter.exe Praxis
call Einrichten.bat starten

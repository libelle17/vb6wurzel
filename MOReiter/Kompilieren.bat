@echo off
rem Kompiliert MOReiter.au3 mit dem Reiter-Icon
cd /d "%~dp0"
set A2E=%ProgramFiles(x86)%\AutoIt3\Aut2Exe\Aut2exe.exe
if not exist "%A2E%" set A2E=%ProgramFiles%\AutoIt3\Aut2Exe\Aut2exe.exe
"%A2E%" /in MOReiter.au3 /out MOReiter.exe /icon MOReiter.ico /x86
if errorlevel 1 (echo Kompilieren fehlgeschlagen & pause)

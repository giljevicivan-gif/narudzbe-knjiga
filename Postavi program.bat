@echo off
rem Jednokratno postavljanje: ikona na radnoj povrsini. Python NE treba instalirati.
cd /d "%~dp0"
if not exist "python\python.exe" (
  echo Nedostaje mapa "python".
  echo Ako ste otvorili ZIP, prvo ga raspakirajte: desni klik na zip - "Izdvoji sve" ^(Extract All^),
  echo pa pokrenite "Postavi program.bat" iz raspakirane mape.
  pause
  exit /b
)
"python\python.exe" postavi.py

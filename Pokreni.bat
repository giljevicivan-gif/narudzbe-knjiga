@echo off
cd /d "%~dp0"
title Narudzbe knjiga - ne zatvarati dok radite
if exist "python\python.exe" (
  "python\python.exe" knjige.py
) else (
  echo Nedostaje mapa "python" u mapi programa.
)
pause

@echo off
rem Objavljuje SAMO stranicu za predbilježbu (port 8766) na internetu preko Tailscale Funnel.
rem Prije toga instalirajte Tailscale (tailscale.com/download) i prijavite se.
set TS="C:\Program Files\Tailscale\tailscale.exe"
if not exist %TS% (
  echo Tailscale nije instaliran. Preuzmite ga s https://tailscale.com/download i prijavite se.
  pause
  exit /b
)
%TS% funnel --bg 8766
echo.
echo Adresa za roditelje je gore (https://....ts.net). Upisite je u Postavke programa.
echo Ako pise da Funnel nije ukljucen, otvorite ponudeni link, ukljucite ga i ponovno pokrenite ovu datoteku.
%TS% funnel status
pause

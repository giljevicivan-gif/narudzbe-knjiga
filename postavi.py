"""Jednokratno postavljanje na računalu: ikona "Narudžbe knjiga" na radnoj površini (+ po želji pokretanje s Windowsima).

Pokreće se preko "Postavi program.bat". Prečac pokušava napraviti na tri načina (VBScript, PowerShell, a ako
ni to ne prođe, na radnu površinu stavi malu .bat datoteku koja pokreće program).
"""
import base64
import ctypes
import os
import subprocess
import sys
import tempfile

MAPA = os.path.dirname(os.path.abspath(__file__))
PYW = os.path.join(MAPA, 'python', 'pythonw.exe')
IKONA = os.path.join(MAPA, 'knjige.ico')
IME = 'Narudžbe knjiga'


def mapa_windows(csidl):
    """Prava mapa radne površine / Startupa (radi i kad je radna površina preusmjerena na OneDrive)."""
    buf = ctypes.create_unicode_buffer(260)
    if ctypes.windll.shell32.SHGetFolderPathW(None, csidl, None, 0, buf) == 0 and os.path.isdir(buf.value):
        return buf.value
    for p in ('~/Desktop', '~/OneDrive/Desktop', '~/OneDrive/Radna površina'):
        if os.path.isdir(os.path.expanduser(p)):
            return os.path.expanduser(p)
    return None


def vbs_precac(lnk, argumenti):
    kod = ('Set w = CreateObject("WScript.Shell")\r\n'
           'Set s = w.CreateShortcut("%s")\r\n'
           's.TargetPath = "%s"\r\ns.Arguments = "%s"\r\ns.WorkingDirectory = "%s"\r\n'
           's.IconLocation = "%s"\r\ns.Description = "%s"\r\ns.Save\r\n') % (lnk, PYW, argumenti, MAPA, IKONA, IME)
    f = os.path.join(tempfile.gettempdir(), 'knjige_precac.vbs')
    with open(f, 'w', encoding='utf-16') as d:  # UTF-16 da prođu i slova č, ć, š, ž u putanji
        d.write(kod)
    r = subprocess.run(['cscript', '//nologo', f], capture_output=True, text=True)
    os.remove(f)
    return r.stdout + r.stderr


def ps_precac(lnk, argumenti):
    nav = lambda s: "'" + s.replace("'", "''") + "'"
    kod = ('$w = New-Object -ComObject WScript.Shell; $s = $w.CreateShortcut(%s); $s.TargetPath = %s; '
           '$s.Arguments = %s; $s.WorkingDirectory = %s; $s.IconLocation = %s; $s.Save()') % (
        nav(lnk), nav(PYW), nav(argumenti), nav(MAPA), nav(IKONA))
    r = subprocess.run(['powershell', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-EncodedCommand',
                        base64.b64encode(kod.encode('utf-16-le')).decode()], capture_output=True, text=True)
    return r.stdout + r.stderr


def napravi(mapa, argumenti):
    lnk = os.path.join(mapa, IME + '.lnk')
    greske = []
    for nacin in (vbs_precac, ps_precac):
        try:
            greske.append(nacin(lnk, argumenti).strip())
        except Exception as e:
            greske.append(str(e))
        if os.path.exists(lnk):
            return lnk
    # zadnja mogućnost: obična .bat datoteka koja pokrene program bez crnog prozora
    bat = os.path.join(mapa, IME.replace('ž', 'z') + '.bat')
    with open(bat, 'w', encoding='utf-8') as f:
        f.write('@echo off\r\nchcp 65001 >nul\r\ncd /d "%s"\r\nstart "" "%s" knjige.py %s\r\n' % (MAPA, PYW, argumenti))
    print('  (prečac nije uspio: %s)' % ' | '.join(g for g in greske if g)[:300])
    return bat


def main():
    print('Postavljanje programa "Narudžbe knjiga"\n')
    temp = tempfile.gettempdir().lower()
    if MAPA.lower().startswith(temp) or not os.path.exists(os.path.join(MAPA, 'knjige.py')):
        print('Program je pokrenut iz ZIP datoteke. Prvo ga raspakirajte:')
        print('desni klik na zip -> "Izdvoji sve" (Extract All), npr. u C:\\Knjige,')
        print('pa tamo ponovno pokrenite "Postavi program.bat".')
        return
    if not os.path.exists(PYW):
        print('Nedostaje python\\pythonw.exe - kopirajte cijelu mapu programa.')
        return
    radna = mapa_windows(0x10)  # CSIDL_DESKTOPDIRECTORY
    if not radna:
        print('Nisam pronašao radnu površinu.')
    else:
        print('Ikona na radnoj površini:', napravi(radna, ''))
    odg = input('\nPokretati program automatski kad se upali računalo (za online predbilježbu)? [D/N]: ')
    startup = mapa_windows(0x07)  # CSIDL_STARTUP
    if startup:
        for ime in (IME + '.lnk', IME.replace('ž', 'z') + '.bat', 'Narudzbe knjiga.lnk'):
            p = os.path.join(startup, ime)
            if os.path.exists(p):
                os.remove(p)
        if odg.strip().lower()[:1] in ('d', 'y'):
            print('Automatsko pokretanje:', napravi(startup, '--bez-preglednika'))
    print('\nGotovo. Program pokrećete ikonom "%s" na radnoj površini.' % IME)
    print('Radi u pozadini; potpuno se zatvara u Postavkama gumbom "Zatvori program".')


if __name__ == '__main__':
    try:
        main()
    except Exception as e:
        print('\nGREŠKA:', e)
    input('\nPritisnite Enter za kraj...')

"""Nadogradnja programa bez ponovnog pokretanja Setupa.

Nadogradnja je zip "Narudzbe-knjiga-nadogradnja-<verzija>.zip" u kojem su samo datoteke programa
(knjige.py, index.html...) i verzija.txt. Podaci (knjige.db, postavke.json, logo.png, kopije) se ne diraju.
Program sam traži takav zip u mapama Preuzimanja i Radna površina, a može se i ručno učitati u Postavkama.
Prije zamjene stare datoteke se spreme u mapu "stara_verzija"; ako nova verzija ne prođe provjeru,
stare se vraćaju.
"""
import glob
import io
import os
import re
import shutil
import subprocess
import sys
import time
import urllib.request
import zipfile

MAPA = os.path.dirname(os.path.abspath(__file__))
STARA = os.path.join(MAPA, 'stara_verzija')
DOZVOLJENE = {'knjige.py', 'javno.py', 'popis.py', 'obrasci.py', 'dbf.py', 'nadogradnja.py', 'postavi.py',
              'index.html', 'javno.html', 'UPUTE.txt', 'knjige.ico', 'pypdf.zip', 'Pokreni.bat',
              'Postavi program.bat', 'Objavi predbiljezbu na internetu.bat'}
UZORAK = 'Narudzbe-knjiga-nadogradnja*.zip'
GITHUB = 'giljevicivan-gif/narudzbe-knjiga'  # javni repozitorij s kodom programa (bez podataka)
_gh = {'vrijeme': 0, 'verzija': None}


def kljuc(verzija):
    return tuple(int(x) for x in re.findall(r'\d+', verzija or ''))


def procitaj(sadrzaj):
    """Provjeri zip nadogradnje; vraća (verzija, {ime: bajtovi})."""
    try:
        z = zipfile.ZipFile(io.BytesIO(sadrzaj))
    except zipfile.BadZipFile:
        raise ValueError('Ovo nije datoteka nadogradnje (.zip).')
    imena = [i.filename for i in z.infolist() if not i.is_dir()]
    if 'verzija.txt' not in imena:
        raise ValueError('Ovo nije nadogradnja programa Narudžbe knjiga (nema verzija.txt).')
    verzija = z.read('verzija.txt').decode('utf-8').strip()
    datoteke = {}
    for ime in imena:
        if ime == 'verzija.txt':
            continue
        if ime not in DOZVOLJENE:
            raise ValueError('Nadogradnja sadrži nepoznatu datoteku: %s' % ime)
        datoteke[ime] = z.read(ime)
    if 'knjige.py' not in datoteke:
        raise ValueError('Nadogradnja nema knjige.py.')
    for ime, d in datoteke.items():
        if ime.endswith('.py'):
            try:
                compile(d, ime, 'exec')
            except SyntaxError as e:
                raise ValueError('Nadogradnja je oštećena (%s: %s).' % (ime, e))
    return verzija, datoteke


def pronadji(trenutna, dodatne_mape=()):
    """Najnoviji zip nadogradnje (noviji od trenutne verzije) u Preuzimanjima / na radnoj površini."""
    dom = os.path.expanduser('~')
    mape = [os.path.join(dom, 'Downloads'), os.path.join(dom, 'Preuzimanja'), os.path.join(dom, 'Desktop'),
            os.path.join(dom, 'OneDrive', 'Desktop'), os.path.join(dom, 'OneDrive', 'Radna površina'), MAPA]
    mape += [m for m in dodatne_mape if m]
    najbolja = None
    for m in mape:
        for put in glob.glob(os.path.join(m, UZORAK)):
            try:
                with open(put, 'rb') as f:
                    verzija, _ = procitaj(f.read())
            except (OSError, ValueError):
                continue
            if kljuc(verzija) > kljuc(trenutna) and (not najbolja or kljuc(verzija) > kljuc(najbolja['verzija'])):
                najbolja = {'put': put, 'verzija': verzija, 'ime': os.path.basename(put)}
    return najbolja


def _provjeri_pokretanje():
    r = subprocess.run([_python(), os.path.join(MAPA, 'knjige.py'), '--provjera'], cwd=MAPA,
                       capture_output=True, text=True, timeout=60, creationflags=_bez_prozora())
    if r.returncode != 0:
        raise ValueError('Nova verzija se ne pokreće: ' + (r.stderr or r.stdout).strip()[-400:])


def _python():
    exe = sys.executable
    if os.name == 'nt' and exe.lower().endswith('python.exe'):
        pyw = exe[:-10] + 'pythonw.exe'
        if os.path.exists(pyw):
            return pyw
    return exe


def _bez_prozora():
    return 0x08000000 if os.name == 'nt' else 0  # CREATE_NO_WINDOW


def instaliraj(sadrzaj):
    """Zamijeni datoteke programa; ako nova verzija ne prođe provjeru, vrati staru. Vraća novu verziju."""
    verzija, datoteke = procitaj(sadrzaj)
    os.makedirs(STARA, exist_ok=True)
    stare = {}
    for ime in datoteke:
        put = os.path.join(MAPA, ime)
        if os.path.exists(put):
            shutil.copy2(put, os.path.join(STARA, ime))
            stare[ime] = True
    try:
        for ime, d in datoteke.items():
            privremena = os.path.join(MAPA, ime + '.novo')
            with open(privremena, 'wb') as f:
                f.write(d)
            os.replace(privremena, os.path.join(MAPA, ime))
        shutil.rmtree(os.path.join(MAPA, '__pycache__'), ignore_errors=True)
        _provjeri_pokretanje()
    except Exception:
        for ime in datoteke:  # vrati staru verziju
            put = os.path.join(MAPA, ime)
            if ime in stare:
                shutil.copy2(os.path.join(STARA, ime), put)
            elif os.path.exists(put):
                os.remove(put)
        raise
    return verzija


def ponovno_pokreni():
    """Pokreni novu verziju (ona pričeka da se ova ugasi)."""
    subprocess.Popen([_python(), os.path.join(MAPA, 'knjige.py'), '--bez-preglednika', '--nakon-nadogradnje'],
                     cwd=MAPA, creationflags=_bez_prozora() | (0x00000008 if os.name == 'nt' else 0),
                     close_fds=True)


def _dohvati(url, timeout=10):
    req = urllib.request.Request(url, headers={'User-Agent': 'narudzbe-knjiga', 'Cache-Control': 'no-cache'})
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return r.read()


def github_verzija(repo=GITHUB):
    """Verzija na GitHubu (provjerava se najviše jednom u 30 minuta; bez interneta vraća None)."""
    if time.time() - _gh['vrijeme'] > 1800:
        _gh['vrijeme'] = time.time()
        try:
            _gh['verzija'] = _dohvati('https://raw.githubusercontent.com/%s/main/verzija.txt?t=%d'
                                      % (repo, time.time()), 5).decode('utf-8').strip()
        except Exception:
            _gh['verzija'] = None
    return _gh['verzija']


def github_zip(repo=GITHUB):
    """Skini kod s GitHuba i složi ga u zip nadogradnje (samo dopuštene datoteke)."""
    z = zipfile.ZipFile(io.BytesIO(_dohvati('https://codeload.github.com/%s/zip/refs/heads/main' % repo, 60)))
    izlaz = io.BytesIO()
    with zipfile.ZipFile(izlaz, 'w', zipfile.ZIP_DEFLATED) as zout:
        for i in z.infolist():
            ime = i.filename.split('/', 1)[-1]  # makni mapu "narudzbe-knjiga-main/"
            if ime in DOZVOLJENE or ime == 'verzija.txt':
                zout.writestr(ime, z.read(i.filename))
    _gh['vrijeme'] = 0
    return izlaz.getvalue()

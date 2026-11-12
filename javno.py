"""Javni dio programa: online predbilježba za roditelje.

Radi na zasebnom portu (zadano 8766) i vidi samo ovo: popis kompleta s cijenama, slanje predbilježbe
i provjeru statusa. Ostatak programa (narudžbe, kupci, postavke) s ove adrese nije dostupan.
"""
import collections
import datetime
import json
import os
import re
import threading
import time
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import parse_qs, urlparse

MAPA = os.path.dirname(os.path.abspath(__file__))
_db = None  # funkcija za otvaranje baze (transakcija) - postavi je knjige.py
_slanja = collections.defaultdict(list)
_brava = threading.Lock()


def _ogranicenje(ip):
    """Najviše 5 slanja po adresi u 10 minuta i 300 ukupno na dan (zaštita od zlouporabe)."""
    sad = time.time()
    with _brava:
        for k in list(_slanja):
            _slanja[k] = [t for t in _slanja[k] if sad - t < 86400]
            if not _slanja[k]:
                del _slanja[k]
        if len([t for t in _slanja[ip] if sad - t < 600]) >= 5:
            return 'Previše slanja u kratkom vremenu. Pokušajte ponovno za 10 minuta.'
        if sum(len(v) for v in _slanja.values()) >= 300:
            return 'Danas je primljeno previše predbilježbi. Nazovite nas telefonom.'
        _slanja[ip].append(sad)
    return None


def _postavke(c):
    return {r['k']: r['v'] for r in c.execute('SELECT k, v FROM postavke')}


def _sezona(c):
    return int(_postavke(c).get('sezona') or datetime.date.today().year)


def podaci(c):
    p = _postavke(c)
    sez = _sezona(c)
    otvoreno = p.get('predbiljezba_ukljucena') == '1'
    rez = {'firma': p.get('firma', ''), 'adresa': p.get('adresa', ''), 'telefon': p.get('telefon', ''),
           'tekst': p.get('predbiljezba_tekst', ''), 'otvoreno': otvoreno, 'sezona': sez}
    if not otvoreno:
        return rez
    ks = [dict(r) for r in c.execute(
        'SELECT k.id, k.naziv, k.razred, k.skola_id, s.naziv AS skola FROM kompleti k JOIN skole s ON s.id = k.skola_id '
        'WHERE COALESCE(k.azurirano, 0) >= ? ORDER BY s.naziv, k.razred, k.naziv', (sez,))]
    knjige = {}
    for r in c.execute('SELECT kk.komplet_id, kn.id, kn.naziv, kn.vrsta, kn.nakladnik, kn.cijena FROM komplet_knjige kk '
                       'JOIN knjige kn ON kn.id = kk.knjiga_id WHERE kn.aktivna = 1 ORDER BY kn.naziv'):
        knjige.setdefault(r[0], []).append({'id': r[1], 'naziv': r[2], 'vrsta': r[3], 'nakladnik': r[4], 'cijena': r[5]})
    rez['kompleti'] = [dict(k, knjige=knjige[k['id']]) for k in ks if k['id'] in knjige]
    return rez


def _tekst(d, k, najvise):
    v = str(d.get(k) or '').strip()
    if len(v) > najvise:
        raise ValueError('Predugačak unos: ' + k)
    return v


def predbiljezba(c, d, ip):
    if d.get('web'):  # skriveno polje - popunjavaju ga samo roboti
        return {'broj': 0}
    if _postavke(c).get('predbiljezba_ukljucena') != '1':
        raise ValueError('Predbilježbe trenutno nisu otvorene.')
    kupac, telefon = _tekst(d, 'kupac', 80), _tekst(d, 'telefon', 30)
    email, napomena = _tekst(d, 'email', 100), _tekst(d, 'napomena', 500)
    if len(kupac) < 3:
        raise ValueError('Upišite ime i prezime.')
    if len(re.sub(r'\D', '', telefon)) < 8:
        raise ValueError('Upišite ispravan broj telefona.')
    if email and not re.fullmatch(r'[^@\s]+@[^@\s]+\.[^@\s]+', email):
        raise ValueError('E-mail adresa nije ispravna.')
    if not d.get('privola'):
        raise ValueError('Potrebna je privola za obradu podataka.')
    k = c.execute('SELECT id, skola_id, razred FROM kompleti WHERE id = ?', (int(d.get('komplet_id') or 0),)).fetchone()
    if not k:
        raise ValueError('Odaberite školu, razred i smjer.')
    dozvoljene = {r[0] for r in c.execute('SELECT knjiga_id FROM komplet_knjige WHERE komplet_id = ?', (k['id'],))}
    knjige = [int(x) for x in d.get('knjige') or [] if int(x) in dozvoljene]
    if not knjige:
        raise ValueError('Označite barem jednu knjigu.')
    # iza tunela sve dolazi s 127.0.0.1, pa se tada broji po telefonu da jedan roditelj ne blokira ostale
    greska = _ogranicenje(ip + re.sub(r'\D', '', telefon)[-8:] if ip in ('127.0.0.1', '::1') else ip)
    if greska:
        raise ValueError(greska)
    sez = _sezona(c)
    broj = c.execute('SELECT COALESCE(MAX(broj), 0) + 1 FROM predbiljezbe WHERE godina = ?', (sez,)).fetchone()[0]
    c.execute('INSERT INTO predbiljezbe(godina, broj, vrijeme, kupac, telefon, email, skola_id, razred, komplet_id, knjige, '
              "napomena, status, ip) VALUES(?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'nova', ?)",
              (sez, broj, datetime.datetime.now().isoformat(timespec='seconds'), kupac, telefon, email, k['skola_id'],
               k['razred'], k['id'], json.dumps(knjige), napomena, ip[:60]))
    return {'broj': broj, 'godina': sez}


def status(c, broj, telefon):
    znam = re.sub(r'\D', '', telefon or '')[-8:]
    if len(znam) < 6 or not broj.isdigit():
        raise ValueError('Upišite broj predbilježbe i telefon.')
    p = c.execute('SELECT p.status, p.telefon, n.status AS nstatus FROM predbiljezbe p '
                  'LEFT JOIN narudzbe n ON n.id = p.narudzba_id WHERE p.godina = ? AND p.broj = ?',
                  (_sezona(c), int(broj))).fetchone()
    if not p or not re.sub(r'\D', '', p['telefon']).endswith(znam):
        raise ValueError('Predbilježba s tim brojem i telefonom nije pronađena.')
    opis = {'nova': 'Zaprimljena, još nije obrađena.', 'odbijeno': 'Nije prihvaćena - nazovite nas.',
            'naručeno': 'Knjige su naručene, javit ćemo Vam kad stignu.', 'stiglo': 'Knjige su stigle! Možete doći po njih.',
            'preuzeto': 'Knjige su preuzete. Hvala!', 'otkazano': 'Narudžba je otkazana.'}
    s = p['nstatus'] or p['status']
    return {'status': opis.get(s, s)}


class JavniHandler(BaseHTTPRequestHandler):
    server_version = 'Predbiljezba'
    sys_version = ''

    def log_message(self, *a):
        pass

    def ip(self):
        return (self.headers.get('Cf-Connecting-Ip') or self.headers.get('X-Forwarded-For') or self.client_address[0]).split(',')[0].strip()

    def posalji(self, kod, tijelo, tip='application/json; charset=utf-8'):
        self.send_response(kod)
        self.send_header('Content-Type', tip)
        self.send_header('Cache-Control', 'no-store')
        self.send_header('X-Content-Type-Options', 'nosniff')
        self.send_header('X-Frame-Options', 'DENY')
        self.send_header('Referrer-Policy', 'no-referrer')
        self.send_header('Content-Security-Policy', "default-src 'self'; style-src 'self' 'unsafe-inline'; "
                                                    "script-src 'self' 'unsafe-inline'; img-src 'self' data:")
        self.send_header('Content-Length', str(len(tijelo)))
        self.end_headers()
        self.wfile.write(tijelo)

    def obradi(self, fn):
        try:
            with _db() as c:
                rez = fn(c)
            self.posalji(200, json.dumps(rez, ensure_ascii=False).encode('utf-8'))
        except ValueError as e:
            self.posalji(400, json.dumps({'greska': str(e)}, ensure_ascii=False).encode('utf-8'))
        except Exception:
            self.posalji(400, '{"greska": "Greška. Pokušajte ponovno ili nas nazovite."}'.encode('utf-8'))

    def do_GET(self):
        u = urlparse(self.path)
        q = {k: v[0] for k, v in parse_qs(u.query).items()}
        if u.path in ('/', '/index.html'):
            with open(os.path.join(MAPA, 'javno.html'), 'rb') as f:
                return self.posalji(200, f.read(), 'text/html; charset=utf-8')
        if u.path == '/api/podaci':
            return self.obradi(podaci)
        if u.path == '/api/status':
            return self.obradi(lambda c: status(c, q.get('broj', ''), q.get('telefon', '')))
        self.posalji(404, b'{}')

    def do_POST(self):
        n = int(self.headers.get('Content-Length') or 0)
        if urlparse(self.path).path != '/api/predbiljezba' or n > 20000:
            return self.posalji(404, b'{}')
        try:
            d = json.loads(self.rfile.read(n) or b'{}')
        except ValueError:
            return self.posalji(400, b'{}')
        self.obradi(lambda c: predbiljezba(c, d, self.ip()))


def pokreni(transakcija, port, adresa='127.0.0.1'):
    global _db
    _db = transakcija
    srv = ThreadingHTTPServer((adresa, port), JavniHandler)
    threading.Thread(target=srv.serve_forever, daemon=True).start()
    return srv

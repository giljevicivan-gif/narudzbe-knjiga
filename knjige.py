"""Narudžbe knjiga - jednostavan program za knjižaru.

Pokretanje:  python knjige.py            (otvara preglednik)
Prijenos iz starog programa:  python knjige.py --uvoz "C:\\put\\do\\RABLJENJE"
"""
import contextlib
import csv
import datetime
import io
import json
import os
import re
import shutil
import sqlite3
import sys
import threading
import time
import webbrowser
import base64
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import parse_qs, urlparse

import dbf
import javno
import nadogradnja
import obrasci
import popis

VERZIJA = '2026.11.13'
MAPA = os.path.dirname(os.path.abspath(__file__))
BAZA = os.path.join(MAPA, 'knjige.db')
CONFIG = os.path.join(MAPA, 'postavke.json')
STATUSI = ['naručeno', 'stiglo', 'preuzeto', 'otkazano']

SHEMA = """
CREATE TABLE IF NOT EXISTS knjige(id INTEGER PRIMARY KEY, naziv TEXT NOT NULL, vrsta TEXT DEFAULT '',
  autori TEXT DEFAULT '', nakladnik TEXT DEFAULT '', cijena REAL DEFAULT 0, barkod TEXT DEFAULT '',
  vulpes_sifra INTEGER, aktivna INTEGER DEFAULT 1, stari_broj INTEGER);
CREATE TABLE IF NOT EXISTS skole(id INTEGER PRIMARY KEY, naziv TEXT UNIQUE NOT NULL);
CREATE TABLE IF NOT EXISTS kompleti(id INTEGER PRIMARY KEY, naziv TEXT NOT NULL, skola_id INTEGER, razred TEXT DEFAULT '');
CREATE TABLE IF NOT EXISTS komplet_knjige(komplet_id INTEGER, knjiga_id INTEGER, PRIMARY KEY(komplet_id, knjiga_id));
CREATE TABLE IF NOT EXISTS narudzbe(id INTEGER PRIMARY KEY, godina INTEGER, broj INTEGER, datum TEXT,
  kupac TEXT DEFAULT '', telefon TEXT DEFAULT '', skola_id INTEGER, razred TEXT DEFAULT '',
  akontacija REAL DEFAULT 0, nacin_pl TEXT DEFAULT '', napomena TEXT DEFAULT '', status TEXT DEFAULT 'naručeno');
CREATE TABLE IF NOT EXISTS stavke(id INTEGER PRIMARY KEY, narudzba_id INTEGER, knjiga_id INTEGER,
  naziv TEXT, barkod TEXT DEFAULT '', nakladnik TEXT DEFAULT '', cijena REAL DEFAULT 0, kolicina INTEGER DEFAULT 1);
CREATE TABLE IF NOT EXISTS postavke(k TEXT PRIMARY KEY, v TEXT);
CREATE INDEX IF NOT EXISTS ix_stavke ON stavke(narudzba_id);
CREATE INDEX IF NOT EXISTS ix_barkod ON knjige(barkod);
CREATE TABLE IF NOT EXISTS obrasci_nakladnika(nakladnik TEXT PRIMARY KEY, ime TEXT, sadrzaj BLOB, datum TEXT);
CREATE TABLE IF NOT EXISTS predbiljezbe(id INTEGER PRIMARY KEY, godina INTEGER, broj INTEGER, vrijeme TEXT, kupac TEXT,
  telefon TEXT, email TEXT, skola_id INTEGER, razred TEXT, komplet_id INTEGER, knjige TEXT, napomena TEXT,
  status TEXT DEFAULT 'nova', narudzba_id INTEGER, ip TEXT);
CREATE TABLE IF NOT EXISTS mzo_katalog(sifra TEXT, naziv TEXT, autori TEXT DEFAULT '', nakladnik TEXT DEFAULT '',
  razred TEXT DEFAULT '', predmet TEXT DEFAULT '', cijena REAL DEFAULT 0);
CREATE TABLE IF NOT EXISTS nakladnici(naziv TEXT PRIMARY KEY, email TEXT DEFAULT '', kontakt TEXT DEFAULT '');
CREATE TABLE IF NOT EXISTS narudzbenice(id INTEGER PRIMARY KEY, godina INTEGER, broj INTEGER, datum TEXT,
  nakladnik TEXT, napomena TEXT DEFAULT '');
CREATE TABLE IF NOT EXISTS narudzbenice_stavke(id INTEGER PRIMARY KEY, narudzbenica_id INTEGER, naziv TEXT,
  barkod TEXT DEFAULT '', kolicina INTEGER, cijena REAL DEFAULT 0);
CREATE TABLE IF NOT EXISTS rabljene(id INTEGER PRIMARY KEY, knjiga_id INTEGER, naziv TEXT, barkod TEXT DEFAULT '',
  ocuvanost TEXT DEFAULT '', otkup_cijena REAL DEFAULT 0, prodajna_cijena REAL DEFAULT 0, datum_otkupa TEXT,
  prodavatelj TEXT DEFAULT '', telefon TEXT DEFAULT '', status TEXT DEFAULT 'na stanju', datum_prodaje TEXT,
  kupac TEXT DEFAULT '');
"""

ZADANE_POSTAVKE = {
    'firma': 'Knjižara PETICA',
    'adresa': 'K. Zvonimira 14, Metković',
    'telefon': '020/681-140',
    'napomena_listic': '',
    'vulpes_artikli': '',
    'sezona': '',
    'poruka_kupcu': 'Poštovani, Vaše knjige (narudžba br. {broj}) su stigle i možete ih preuzeti u {firma}. '
                    'Za platiti: {ostatak}. Hvala!',
    'rabljene_otkup_posto': '30',
    'rabljene_prodaja_posto': '50',
    'poruka_podsjetnik': 'Poštovani, podsjećamo da Vas knjige (narudžba br. {broj}) čekaju u {firma}. '
                         'Za platiti: {ostatak}. Hvala!',
    'podsjetnik_dana': '7',
    'predbiljezba_ukljucena': '0',
    'predbiljezba_tekst': 'Odaberite školu i razred, označite knjige koje trebate i pošaljite predbilježbu. '
                          'Javit ćemo Vam se SMS-om ili WhatsAppom kad knjige stignu. Plaćanje pri preuzimanju.',
    'predbiljezba_adresa': '',
    'kopija_mapa': '',
    'kopija_stanje': '',
}
# ključ knjige u stavkama: barkod ako postoji, inače naziv
KLJUC = "COALESCE(NULLIF(TRIM({0}barkod), ''), {0}naziv)"
KOPIJE = os.path.join(MAPA, 'sigurnosne_kopije')


def db():
    c = sqlite3.connect(BAZA, timeout=10)
    c.row_factory = sqlite3.Row
    c.execute('PRAGMA journal_mode=WAL')
    return c


@contextlib.contextmanager
def transakcija():
    c = db()
    try:
        with c:
            yield c
    finally:
        c.close()


def init_db():
    with transakcija() as c:
        c.executescript(SHEMA)
        if 'sifra_mzo' not in [r[1] for r in c.execute('PRAGMA table_info(knjige)')]:
            c.execute("ALTER TABLE knjige ADD COLUMN sifra_mzo TEXT DEFAULT ''")
        if 'stiglo' not in [r[1] for r in c.execute('PRAGMA table_info(stavke)')]:
            c.execute('ALTER TABLE stavke ADD COLUMN stiglo INTEGER DEFAULT 0')
        if 'predmet' not in [r[1] for r in c.execute('PRAGMA table_info(komplet_knjige)')]:
            c.execute("ALTER TABLE komplet_knjige ADD COLUMN predmet TEXT DEFAULT ''")
        stupci_n = [r[1] for r in c.execute('PRAGMA table_info(narudzbe)')]
        if 'platitelj' not in stupci_n:
            c.execute("ALTER TABLE narudzbe ADD COLUMN platitelj TEXT DEFAULT ''")
        if 'stiglo_datum' not in stupci_n:
            c.execute('ALTER TABLE narudzbe ADD COLUMN stiglo_datum TEXT')
        if 'tip' not in [r[1] for r in c.execute('PRAGMA table_info(skole)')]:
            c.execute('ALTER TABLE skole ADD COLUMN tip TEXT')
        if 'razina' not in [r[1] for r in c.execute('PRAGMA table_info(narudzbenice)')]:
            c.execute("ALTER TABLE narudzbenice ADD COLUMN razina TEXT DEFAULT ''")
        if 'tip' not in [r[1] for r in c.execute('PRAGMA table_info(narudzbenice)')]:
            c.execute("ALTER TABLE narudzbenice ADD COLUMN tip TEXT DEFAULT 'narudzba'")
        nova_kolona = 'azurirano' not in [r[1] for r in c.execute('PRAGMA table_info(kompleti)')]
        if nova_kolona:
            c.execute('ALTER TABLE kompleti ADD COLUMN azurirano INTEGER')
        for k, v in ZADANE_POSTAVKE.items():
            c.execute('INSERT OR IGNORE INTO postavke(k, v) VALUES(?, ?)', (k, v))
        if not postavke(c).get('sezona'):
            g = c.execute('SELECT MAX(godina) FROM narudzbe').fetchone()[0] or datetime.date.today().year
            c.execute("UPDATE postavke SET v = ? WHERE k = 'sezona'", (str(g),))
        if nova_kolona:  # postojeći kompleti vrijede za tekuću sezonu
            c.execute('UPDATE kompleti SET azurirano = ?', (sezona(c),))


def sezona(c):
    try:
        return int(postavke(c).get('sezona'))
    except (TypeError, ValueError):
        return datetime.date.today().year


def sigurnosna_kopija(naziv=None):
    os.makedirs(KOPIJE, exist_ok=True)
    put = os.path.join(KOPIJE, naziv or 'knjige_%s.db' % datetime.datetime.now().strftime('%Y-%m-%d_%H%M'))
    izvor, cilj = sqlite3.connect(BAZA), sqlite3.connect(put)
    try:
        izvor.backup(cilj)
    finally:
        izvor.close()
        cilj.close()
    stare = sorted(f for f in os.listdir(KOPIJE) if f.startswith('dnevna_'))
    for f in stare[:-30]:  # čuvamo zadnjih 30 dnevnih kopija
        os.remove(os.path.join(KOPIJE, f))
    vanjska_kopija(put)
    return put


def vanjska_kopija(put):
    """Kopija i na USB / Google Drive mapu iz Postavki (čuva zadnjih 20). Rezultat se zapiše u postavke."""
    with transakcija() as c:
        mapa = (postavke(c).get('kopija_mapa') or '').strip()
    if not mapa:
        return
    try:
        if not os.path.isdir(mapa):
            raise OSError('mapa "%s" nije dostupna (USB nije utaknut?)' % mapa)
        cilj = os.path.join(mapa, 'knjige_kopija_%s.db' % datetime.datetime.now().strftime('%Y-%m-%d_%H%M'))
        shutil.copyfile(put, cilj)
        stare = sorted(f for f in os.listdir(mapa) if f.startswith('knjige_kopija_'))
        for f in stare[:-20]:
            os.remove(os.path.join(mapa, f))
        stanje = 'OK %s → %s' % (datetime.datetime.now().strftime('%d.%m.%Y. %H:%M'), cilj)
    except OSError as e:
        stanje = 'GREŠKA %s: %s' % (datetime.datetime.now().strftime('%d.%m.%Y. %H:%M'), e)
    with transakcija() as c:
        c.execute('INSERT OR REPLACE INTO postavke(k, v) VALUES(?, ?)', ('kopija_stanje', stanje))


def kopije_u_pozadini():
    """Program može raditi danima (pokreće se s Windowsima) - jednom dnevno napravi kopiju."""
    while True:
        try:
            dnevna = 'dnevna_%s.db' % datetime.date.today().isoformat()
            if not os.path.exists(os.path.join(KOPIJE, dnevna)):
                sigurnosna_kopija(dnevna)
        except Exception as e:
            print('Sigurnosna kopija nije uspjela:', e)
        time.sleep(3600)


def novi_komplet_popis(c, kid):
    c.execute('DELETE FROM komplet_knjige WHERE komplet_id = ?', (kid,))
    c.execute('UPDATE kompleti SET azurirano = ? WHERE id = ?', (sezona(c), kid))


def rows(cur):
    return [dict(r) for r in cur.fetchall()]


def postavke(c):
    return {r['k']: r['v'] for r in c.execute('SELECT k, v FROM postavke')}


def normaliziraj_skolu(naziv):
    s = re.sub(r'\s+', ' ', (naziv or '').upper().replace('.', '. ')).strip()
    s = re.sub(r'\bS\. ?Š\. ?', 'SŠ ', s)
    s = re.sub(r'^SREDNJA (ŠKOLA )?', 'SŠ ', s).replace('GIMAZIJA', 'GIMNAZIJA')
    return re.sub(r'\s+', ' ', s.replace('. ', ' ')).strip()


def razina_skole(naziv):
    """OŠ za osnovne škole, SŠ za sve ostale (srednje, gimnazije...)."""
    return 'OŠ' if re.search(r'(?<!\w)O\.?\s?Š(?!\w)|OSNOVN|\bOS\b', (naziv or '').upper()) else 'SŠ'


def skola_id(c, naziv):
    naziv = normaliziraj_skolu(naziv)
    if not naziv:
        return None
    r = c.execute('SELECT id FROM skole WHERE naziv = ?', (naziv,)).fetchone()
    if r:
        return r['id']
    return c.execute('INSERT INTO skole(naziv, tip) VALUES(?, ?)', (naziv, razina_skole(naziv))).lastrowid


# ---------------------------------------------------------------- Vulpes

_vulpes = {'put': None, 'mtime': None, 'artikli': []}
_vulpes_lock = threading.Lock()


def vulpes_put(c):
    put = (postavke(c).get('vulpes_artikli') or '').strip().strip('"')
    if put and os.path.isdir(put):
        for ime in ('artikli.DBF', 'artikli.dbf', 'ARTIKLI.DBF'):
            if os.path.exists(os.path.join(put, ime)):
                return os.path.join(put, ime)
        return os.path.join(put, 'artikli.DBF')
    return put


def vulpes_artikli(c):
    put = vulpes_put(c)
    if not put:
        raise ValueError('U Postavkama nije upisano gdje je Vulpes (mapa godine, npr. C:\\ROBA\\2026).')
    if not os.path.exists(put):
        raise ValueError('Ne mogu pronaći Vulpes artikle: ' + put)
    with _vulpes_lock:
        mtime = os.path.getmtime(put)
        if _vulpes['put'] != put or _vulpes['mtime'] != mtime:
            art = []
            for r in dbf.citaj(put):
                if r.get('NEAKTIVAN'):
                    continue
                art.append({'sifra': r.get('SIFRA'), 'naziv': r.get('NAZIV', ''),
                            'barkod': (r.get('BARCODE') or '').strip(), 'cijena': r.get('CIJENA') or 0,
                            'stanje': r.get('KOLICINA') or 0})
            _vulpes.update(put=put, mtime=mtime, artikli=art)
        return _vulpes['artikli']


def vulpes_trazi(c, q):
    q = (q or '').strip().upper()
    if not q:
        return []
    art = vulpes_artikli(c)
    if q.isdigit() and len(q) >= 8:
        return [a for a in art if a['barkod'] == q][:50]
    rijeci = q.split()
    return [a for a in art if all(w in a['naziv'].upper() or w == a['barkod'] for w in rijeci)][:80]


PRESKOCI = {'windows', 'program files', 'program files (x86)', 'programdata', 'appdata', '$recycle.bin',
            'system volume information', 'node_modules', 'xml', 'exml', '_backup', 'recovery', 'perflogs'}


def vulpes_pronadi(dubina=4, rok=15):
    """Traži Vulpes mape (artikli.DBF + data1.DBC) po diskovima; vraća najnovije prve."""
    kraj = time.time() + rok
    nadeno = []

    def setaj(mapa, d):
        if d > dubina or time.time() > kraj:
            return
        try:
            imena = os.listdir(mapa)
        except OSError:
            return
        mala = {i.lower() for i in imena}
        if 'artikli.dbf' in mala and 'data1.dbc' in mala:
            put = os.path.join(mapa, next(i for i in imena if i.lower() == 'artikli.dbf'))
            nadeno.append({'mapa': mapa, 'izmijenjeno': datetime.datetime.fromtimestamp(
                os.path.getmtime(put)).strftime('%d.%m.%Y. %H:%M'), '_t': os.path.getmtime(put)})
        for i in imena:
            if i.lower() in PRESKOCI or i.startswith('.'):
                continue
            p = os.path.join(mapa, i)
            if os.path.isdir(p):
                setaj(p, d + 1)

    if os.name == 'nt':
        korijeni = ['%s:\\' % z for z in 'CDEFGHIJKLMNOPQRSTUVWXYZ' if os.path.exists('%s:\\' % z)]
    else:
        korijeni = [os.path.expanduser('~')]
    for k in korijeni:
        setaj(k, 0)
    nadeno.sort(key=lambda x: -x.pop('_t'))
    return nadeno


def vulpes_usporedba(c):
    po_barkodu = {}
    for a in vulpes_artikli(c):
        if a['barkod']:
            po_barkodu.setdefault(a['barkod'], a)
    razlike, nema, povezano = [], [], 0
    for k in rows(c.execute('SELECT * FROM knjige WHERE aktivna = 1 ORDER BY naziv')):
        v = po_barkodu.get((k['barkod'] or '').strip())
        if not v:
            nema.append(k)
            continue
        povezano += 1
        if abs((k['cijena'] or 0) - v['cijena']) > 0.004:
            razlike.append({'knjiga': k, 'vulpes': v})
    return {'razlike': razlike, 'nema': nema, 'povezano': povezano}


# ---------------------------------------------------------------- API

def api_get(c, put, q):
    g = lambda k, d='': q.get(k, [d])[0]
    if put == '/api/nadogradnja':
        dostupna = nadogradnja.pronadji(VERZIJA, [postavke(c).get('kopija_mapa', '')])
        gh = nadogradnja.github_verzija()
        if gh and nadogradnja.kljuc(gh) > nadogradnja.kljuc(dostupna['verzija'] if dostupna else VERZIJA):
            dostupna = {'verzija': gh, 'ime': 's interneta', 'github': True}
        return {'verzija': VERZIJA, 'dostupna': dostupna}
    if put == '/api/init':
        for r in rows(c.execute('SELECT id, naziv FROM skole WHERE tip IS NULL')):
            c.execute('UPDATE skole SET tip = ? WHERE id = ?', (razina_skole(r['naziv']), r['id']))
        return {'skole': rows(c.execute('SELECT * FROM skole ORDER BY naziv')),
                'postavke': postavke(c), 'statusi': STATUSI,
                'godine': sorted({r[0] for r in c.execute('SELECT DISTINCT godina FROM narudzbe')} | {sezona(c)}, reverse=True),
                'sezona': sezona(c), 'verzija': VERZIJA,
                'predbiljezbe_nove': c.execute("SELECT COUNT(*) FROM predbiljezbe WHERE status = 'nova'").fetchone()[0],
                'platitelji': [r[0] for r in c.execute("SELECT DISTINCT platitelj FROM narudzbe WHERE platitelj != '' ORDER BY 1")],
                'otvorene': c.execute("SELECT COUNT(*) FROM narudzbe WHERE status != 'preuzeto' AND godina < ?",
                                      (sezona(c),)).fetchone()[0],
                'kompleti_stari': c.execute('SELECT COUNT(*) FROM kompleti WHERE COALESCE(azurirano, 0) < ?',
                                            (sezona(c),)).fetchone()[0]}
    if put == '/api/knjige':
        return rows(c.execute('SELECT * FROM knjige ORDER BY aktivna DESC, naziv'))
    if put == '/api/kompleti':
        ks = rows(c.execute('SELECT k.*, s.naziv AS skola FROM kompleti k LEFT JOIN skole s ON s.id = k.skola_id '
                            'ORDER BY s.naziv, k.razred, k.naziv'))
        veze = {}
        for r in c.execute('SELECT komplet_id, knjiga_id FROM komplet_knjige'):
            veze.setdefault(r[0], []).append(r[1])
        for k in ks:
            k['knjige'] = veze.get(k['id'], [])
        return ks
    if put == '/api/narudzbe':
        return lista_narudzbi(c, g('godina'), g('status'), g('skola'), g('platitelj', '*'), g('razina'))
    if put == '/api/narudzba':
        n = c.execute('SELECT n.*, s.naziv AS skola FROM narudzbe n LEFT JOIN skole s ON s.id = n.skola_id '
                      'WHERE n.id = ?', (int(g('id')),)).fetchone()
        if not n:
            raise ValueError('Narudžba ne postoji.')
        n = dict(n)
        n['stavke'] = rows(c.execute(
            "SELECT st.*, (SELECT kk.predmet FROM komplet_knjige kk WHERE kk.knjiga_id = st.knjiga_id AND kk.predmet != '' "
            'LIMIT 1) AS predmet FROM stavke st WHERE st.narudzba_id = ? ORDER BY st.id', (n['id'],)))
        return n
    if put == '/api/zbirno':
        return zbirno(c, g('godina'), g('status'), g('skola'), g('razina'))
    if put == '/api/vulpes/trazi':
        return vulpes_trazi(c, g('q'))
    if put == '/api/vulpes/pronadi':
        return vulpes_pronadi()
    if put == '/api/vulpes/usporedba':
        return vulpes_usporedba(c)
    if put == '/api/narudzbenica/prijedlog':
        return prijedlog_narudzbenice(c, g('nakladnik'), g('razina'))
    if put == '/api/povrat/prijedlog':
        return prijedlog_povrata(c, g('nakladnik'))
    if put == '/api/kupac':
        return kupac_po_telefonu(c, g('telefon'))
    if put == '/api/nepreuzeto':
        dana = int(g('dana') or postavke(c).get('podsjetnik_dana') or 7)
        granica = (datetime.date.today() - datetime.timedelta(days=dana)).isoformat()
        return [n for n in lista_narudzbi(c, '', 'stiglo', '') if (n['stiglo_datum'] or n['datum']) <= granica]
    if put == '/api/predbiljezbe':
        st = g('status')
        l = rows(c.execute('SELECT p.*, s.naziv AS skola, k.naziv AS komplet, n.broj AS narudzba_broj FROM predbiljezbe p '
                           'LEFT JOIN skole s ON s.id = p.skola_id LEFT JOIN kompleti k ON k.id = p.komplet_id '
                           'LEFT JOIN narudzbe n ON n.id = p.narudzba_id' + (' WHERE p.status = ?' if st else '') +
                           ' ORDER BY p.id DESC LIMIT 500', (st,) if st else ()))
        for p in l:
            ids = json.loads(p['knjige'] or '[]')
            p['knjige'] = rows(c.execute('SELECT id, naziv, nakladnik, cijena, barkod FROM knjige WHERE id IN (%s)'
                                         % ','.join('?' * len(ids)), ids)) if ids else []
        return l
    if put == '/api/obrasci':  # obrasci za narudžbu po kompletu (knjige redom kao na popisu škole)
        out = []
        for kid in [int(x) for x in g('ids').split(',') if x.strip().isdigit()]:
            k = c.execute('SELECT k.*, s.naziv AS skola FROM kompleti k LEFT JOIN skole s ON s.id = k.skola_id WHERE k.id = ?',
                          (kid,)).fetchone()
            if k:
                out.append(dict(k, knjige=rows(c.execute(
                    'SELECT kn.id, kn.naziv, kn.autori, kn.vrsta, kn.nakladnik, kn.cijena, kk.predmet FROM komplet_knjige kk '
                    'JOIN knjige kn ON kn.id = kk.knjiga_id WHERE kk.komplet_id = ? ORDER BY kk.rowid', (kid,)))))
        return out
    if put == '/api/platitelj':  # zbirni popis za Grad / Županiju / CZSS
        ns = lista_narudzbi(c, g('godina'), g('status'), g('skola'), g('platitelj'), g('razina'))
        for n in ns:
            n['stavke'] = rows(c.execute('SELECT naziv, nakladnik, cijena, kolicina FROM stavke WHERE narudzba_id = ? ORDER BY id',
                                         (n['id'],)))
        return [n for n in ns if n['status'] != 'otkazano']
    if put == '/api/mzo':
        q_ = '%' + g('q') + '%'
        return rows(c.execute('SELECT * FROM mzo_katalog WHERE sifra = ? OR naziv LIKE ? OR autori LIKE ? LIMIT 100',
                              (g('q'), q_, q_)))
    if put == '/api/narudzbenice':
        return rows(c.execute('SELECT n.*, COALESCE(SUM(s.kolicina), 0) AS komada, COUNT(s.id) AS naslova, '
                              'COALESCE(SUM(s.kolicina * s.cijena), 0) AS iznos FROM narudzbenice n '
                              'LEFT JOIN narudzbenice_stavke s ON s.narudzbenica_id = n.id '
                              'GROUP BY n.id ORDER BY n.godina DESC, n.broj DESC'))
    if put == '/api/narudzbenica':
        n = dict(c.execute('SELECT * FROM narudzbenice WHERE id = ?', (int(g('id')),)).fetchone())
        n['stavke'] = rows(c.execute('SELECT * FROM narudzbenice_stavke WHERE narudzbenica_id = ? ORDER BY naziv', (n['id'],)))
        r = c.execute('SELECT * FROM nakladnici WHERE naziv = ?', (n['nakladnik'],)).fetchone()
        n['email'] = r['email'] if r else ''
        o = c.execute('SELECT ime, datum FROM obrasci_nakladnika WHERE nakladnik = ?', (n['nakladnik'],)).fetchone()
        n['obrazac'] = dict(o) if o else None
        return n
    if put == '/api/nakladnici':
        return rows(c.execute('SELECT * FROM nakladnici ORDER BY naziv'))
    if put == '/api/rabljene':
        st = g('status')
        return rows(c.execute('SELECT * FROM rabljene' + (' WHERE status = ?' if st else '') + ' ORDER BY id DESC',
                              (st,) if st else ()))
    if put == '/api/statistika':
        return statistika(c, int(g('godina') or sezona(c)))
    raise KeyError(put)


def filter_sql(godina, status, skola, platitelj=None, razina=''):
    uv, par = [], []
    if razina:
        uv.append('n.skola_id IN (SELECT id FROM skole WHERE tip = ?)'); par.append(razina)
    if platitelj is not None and platitelj != '*':
        uv.append("COALESCE(n.platitelj, '') = ?"); par.append(platitelj)
    if godina:
        uv.append('n.godina = ?'); par.append(int(godina))
    if status:
        uv.append('n.status = ?'); par.append(status)
    if skola:
        uv.append('n.skola_id = ?'); par.append(int(skola))
    return (' WHERE ' + ' AND '.join(uv)) if uv else '', par


def lista_narudzbi(c, godina, status, skola, platitelj=None, razina=''):
    w, par = filter_sql(godina, status, skola, platitelj, razina)
    return rows(c.execute(
        'SELECT n.*, s.naziv AS skola, COALESCE(SUM(st.cijena * st.kolicina), 0) AS iznos, '
        'COALESCE(SUM(st.kolicina), 0) AS komada FROM narudzbe n LEFT JOIN skole s ON s.id = n.skola_id '
        'LEFT JOIN stavke st ON st.narudzba_id = n.id' + w + ' GROUP BY n.id ORDER BY n.godina DESC, n.broj DESC', par))


def zbirno(c, godina, status, skola, razina=''):
    w, par = filter_sql(godina, status, skola, razina=razina)
    if not status:
        w += (' AND' if w else ' WHERE') + " n.status != 'otkazano'"
    return rows(c.execute(
        'SELECT st.naziv, st.barkod, st.nakladnik, st.cijena, SUM(st.kolicina) AS kolicina, '
        'COUNT(DISTINCT n.id) AS narudzbi FROM stavke st JOIN narudzbe n ON n.id = st.narudzba_id' + w +
        ' GROUP BY st.naziv, st.barkod, st.nakladnik ORDER BY st.nakladnik, st.naziv', par))


def api_post(c, put, d):
    if put == '/api/knjiga':
        polja = ('naziv', 'vrsta', 'autori', 'nakladnik', 'cijena', 'barkod', 'vulpes_sifra', 'aktivna')
        vr = [d.get(p) for p in polja]
        if not (d.get('naziv') or '').strip():
            raise ValueError('Upiši naziv knjige.')
        if d.get('id'):
            c.execute('UPDATE knjige SET ' + ', '.join(p + ' = ?' for p in polja) + ' WHERE id = ?', vr + [d['id']])
            return {'id': d['id']}
        return {'id': c.execute('INSERT INTO knjige(' + ', '.join(polja) + ') VALUES(' + ', '.join('?' * len(polja)) + ')',
                                vr).lastrowid}
    if put == '/api/skola':
        if d.get('id') and d.get('tip') in ('OŠ', 'SŠ'):
            c.execute('UPDATE skole SET tip = ? WHERE id = ?', (d['tip'], d['id']))
            return {'id': d['id']}
        return {'id': skola_id(c, d['naziv'])}
    if put == '/api/komplet':
        if not (d.get('naziv') or '').strip():
            raise ValueError('Upiši naziv kompleta (npr. smjer).')
        vr = (d['naziv'].strip(), d.get('skola_id'), d.get('razred', ''))
        if d.get('id'):
            kid = d['id']
            c.execute('UPDATE kompleti SET naziv = ?, skola_id = ?, razred = ? WHERE id = ?', vr + (kid,))
        else:
            kid = c.execute('INSERT INTO kompleti(naziv, skola_id, razred) VALUES(?, ?, ?)', vr).lastrowid
        predmeti = dict(c.execute('SELECT knjiga_id, predmet FROM komplet_knjige WHERE komplet_id = ?', (kid,)).fetchall())
        novi_komplet_popis(c, kid)
        c.executemany('INSERT OR IGNORE INTO komplet_knjige(komplet_id, knjiga_id, predmet) VALUES(?, ?, ?)',
                      [(kid, k, predmeti.get(k, '')) for k in d.get('knjige', [])])
        return {'id': kid}
    if put == '/api/komplet/obrisi':
        c.execute('DELETE FROM komplet_knjige WHERE komplet_id = ?', (d['id'],))
        c.execute('DELETE FROM kompleti WHERE id = ?', (d['id'],))
        return {}
    if put == '/api/narudzba':
        return spremi_narudzbu(c, d)
    if put == '/api/predbiljezba':
        p = c.execute('SELECT * FROM predbiljezbe WHERE id = ?', (d['id'],)).fetchone()
        if d.get('akcija') == 'odbij':
            c.execute("UPDATE predbiljezbe SET status = 'odbijeno' WHERE id = ?", (p['id'],))
            return {}
        if p['narudzba_id']:
            return {'narudzba_id': p['narudzba_id']}
        ids = json.loads(p['knjige'] or '[]')
        knjige = rows(c.execute('SELECT * FROM knjige WHERE id IN (%s)' % ','.join('?' * len(ids)), ids)) if ids else []
        napomena = 'Online predbilježba br. %d' % p['broj'] + ('; ' + p['email'] if p['email'] else '') + \
            ('; ' + p['napomena'] if p['napomena'] else '')
        nid = spremi_narudzbu(c, {'kupac': p['kupac'], 'telefon': p['telefon'], 'skola_id': p['skola_id'],
                                  'razred': p['razred'], 'napomena': napomena,
                                  'stavke': [{'knjiga_id': k['id'], 'naziv': k['naziv'], 'barkod': k['barkod'],
                                              'nakladnik': k['nakladnik'], 'cijena': k['cijena'], 'kolicina': 1}
                                             for k in knjige]})['id']
        c.execute("UPDATE predbiljezbe SET status = 'prihvaćeno', narudzba_id = ? WHERE id = ?", (nid, p['id']))
        return {'narudzba_id': nid}
    if put == '/api/ugasi':
        threading.Thread(target=lambda: (time.sleep(0.5), Handler.server_ref.shutdown()), daemon=True).start()
        return {}
    if put == '/api/nadogradnja':
        if d.get('sadrzaj'):
            sadrzaj = base64.b64decode(d['sadrzaj'])
        elif d.get('github'):
            try:
                sadrzaj = nadogradnja.github_zip()
            except Exception as e:
                raise ValueError('Nadogradnja s interneta nije uspjela (%s). Provjerite internet.' % e)
        else:
            nadjena = nadogradnja.pronadji(VERZIJA, [postavke(c).get('kopija_mapa', '')])
            if not nadjena:
                raise ValueError('Nema novije nadogradnje.')
            with open(nadjena['put'], 'rb') as f:
                sadrzaj = f.read()
        verzija, _ = nadogradnja.procitaj(sadrzaj)
        if nadogradnja.kljuc(verzija) < nadogradnja.kljuc(VERZIJA):
            raise ValueError('To je starija verzija (%s) od ove koja radi (%s).' % (verzija, VERZIJA))
        nadogradnja.instaliraj(sadrzaj)

        def restart():
            time.sleep(0.5)
            nadogradnja.ponovno_pokreni()
            Handler.server_ref.shutdown()
        threading.Thread(target=restart, daemon=True).start()
        return {'verzija': verzija}
    if put == '/api/narudzba/obrisi':
        c.execute('DELETE FROM stavke WHERE narudzba_id = ?', (d['id'],))
        c.execute('DELETE FROM narudzbe WHERE id = ?', (d['id'],))
        return {}
    if put == '/api/narudzbe/status':
        if d['status'] not in STATUSI:
            raise ValueError('Nepoznat status.')
        c.executemany('UPDATE narudzbe SET status = ? WHERE id = ?', [(d['status'], i) for i in d['ids']])
        if d['status'] == 'stiglo':
            c.executemany('UPDATE narudzbe SET stiglo_datum = ? WHERE id = ? AND stiglo_datum IS NULL',
                          [(datetime.date.today().isoformat(), i) for i in d['ids']])
        return {}
    if put == '/api/nova-godina':
        nova = int(d['sezona'])
        if nova <= sezona(c):
            raise ValueError('Nova sezona mora biti veća od trenutne (%d).' % sezona(c))
        kopija = sigurnosna_kopija('prije_sezone_%d.db' % nova)
        c.execute("UPDATE postavke SET v = ? WHERE k = 'sezona'", (str(nova),))
        sakriveno = 0
        if d.get('sakrij_knjige'):
            sakriveno = c.execute('UPDATE knjige SET aktivna = 0 WHERE aktivna = 1 AND id NOT IN '
                                  '(SELECT knjiga_id FROM komplet_knjige)').rowcount
        return {'kopija': kopija, 'sakriveno': sakriveno}
    if put == '/api/kopija':
        return {'kopija': sigurnosna_kopija()}
    if put == '/api/postavke':
        for k, v in d.items():
            if k in ZADANE_POSTAVKE:
                c.execute('INSERT OR REPLACE INTO postavke(k, v) VALUES(?, ?)', (k, v))
        return postavke(c)
    if put in ('/api/popis/analiziraj', '/api/popis/predlozi'):
        katalog = rows(c.execute('SELECT id, naziv, cijena, nakladnik, sifra_mzo FROM knjige'))
        try:
            vulpes = vulpes_artikli(c)
        except ValueError:
            vulpes = []
        if put == '/api/popis/predlozi':
            st = {k: d.get(k) or '' for k in ('naslov', 'autori', 'nakladnik', 'sifra')}
            return popis.predlozi(st, *popis.kandidati(katalog, vulpes))
        redovi = popis.redovi_iz_datoteke(d['ime'], base64.b64decode(d['sadrzaj']))
        sekcije = popis.analiziraj(redovi, katalog, vulpes)
        mzo = {r['sifra']: r for r in rows(c.execute("SELECT * FROM mzo_katalog WHERE sifra != ''"))}
        if mzo:
            kand, po_sifri = popis.kandidati(katalog, vulpes)
            for sek in sekcije:
                for st in sek['stavke']:
                    m = mzo.get(st['sifra'])
                    if not m or st['sifra'] in po_sifri:
                        continue
                    st['mzo'] = m  # službeni naziv iz kataloga ministarstva
                    st['autori'] = st['autori'] or m['autori']
                    st['nakladnik'] = st['nakladnik'] or m['nakladnik']
                    if not st['prijedlozi'] or st['prijedlozi'][0]['ocjena'] < 0.6:
                        st['prijedlozi'] = popis.predlozi(dict(st, naslov=m['naziv']), kand, po_sifri) or st['prijedlozi']
        return {'sekcije': sekcije, 'vulpes': bool(vulpes)}
    if put == '/api/popis/spremi':
        return spremi_popis(c, d)
    if put == '/api/uvoz/mzo':
        return uvoz_mzo(c, d['ime'], base64.b64decode(d['sadrzaj']))
    if put == '/api/uvoz/kompleti':
        return uvoz_kompleta(c, d['ime'], base64.b64decode(d['sadrzaj']))
    if put == '/api/vulpes/preuzmi':
        # preuzmi cijenu (i šifru) iz Vulpesa za odabrane knjige, po barkodu
        po_barkodu = {a['barkod']: a for a in vulpes_artikli(c) if a['barkod']}
        n = 0
        for kid in d['ids']:
            k = c.execute('SELECT barkod FROM knjige WHERE id = ?', (kid,)).fetchone()
            v = k and po_barkodu.get((k['barkod'] or '').strip())
            if v:
                c.execute('UPDATE knjige SET cijena = ?, vulpes_sifra = ? WHERE id = ?', (v['cijena'], v['sifra'], kid))
                n += 1
        return {'azurirano': n}
    if put == '/api/narudzbenica':
        stavke = [s for s in d['stavke'] if int(s.get('kolicina') or 0) > 0]
        if not stavke:
            raise ValueError('Narudžbenica nema nijednu knjigu s količinom.')
        god = sezona(c)
        tip = 'povrat' if d.get('tip') == 'povrat' else 'narudzba'
        broj = c.execute('SELECT COALESCE(MAX(broj), 0) + 1 FROM narudzbenice WHERE godina = ? AND tip = ?', (god, tip)).fetchone()[0]
        nid = c.execute('INSERT INTO narudzbenice(godina, broj, datum, nakladnik, napomena, tip, razina) VALUES(?, ?, ?, ?, ?, ?, ?)',
                        (god, broj, datetime.date.today().isoformat(), d['nakladnik'], d.get('napomena', ''), tip,
                         d.get('razina') if d.get('razina') in ('OŠ', 'SŠ') else '')).lastrowid
        c.executemany('INSERT INTO narudzbenice_stavke(narudzbenica_id, naziv, barkod, kolicina, cijena) VALUES(?, ?, ?, ?, ?)',
                      [(nid, s['naziv'], s.get('barkod', ''), int(s['kolicina']), float(s.get('cijena') or 0)) for s in stavke])
        return {'id': nid}
    if put == '/api/narudzbenice/sve':  # jednim klikom narudžbenice za sve nakladnike s predloženim količinama
        razina = d.get('razina') or ''
        ids = []
        for s in prijedlog_narudzbenice(c, '', razina):
            if s['komada'] and s['nakladnik'] != 'bez nakladnika':
                st = [r for r in prijedlog_narudzbenice(c, s['nakladnik'], razina) if r['kolicina']]
                ids.append(api_post(c, '/api/narudzbenica', {'nakladnik': s['nakladnik'], 'stavke': st, 'razina': razina,
                                                             'napomena': d.get('napomena', '')})['id'])
        return {'ids': ids}
    if put == '/api/narudzbenica/excel':  # popuni Excel narudžbenicu nakladnika količinama iz ove narudžbenice
        n = api_get(c, '/api/narudzbenica', {'id': [str(d['id'])]})
        if d.get('sadrzaj'):  # novi obrazac - zapamti ga za tog nakladnika
            obrasci.popuni(base64.b64decode(d['sadrzaj']), {})  # provjera da je ispravan
            c.execute('INSERT OR REPLACE INTO obrasci_nakladnika VALUES(?, ?, ?, ?)',
                      (n['nakladnik'], d['ime'], base64.b64decode(d['sadrzaj']), datetime.date.today().isoformat()))
        o = c.execute('SELECT * FROM obrasci_nakladnika WHERE nakladnik = ?', (n['nakladnik'],)).fetchone()
        if not o:
            raise ValueError('Za nakladnika %s još nije učitan Excel obrazac.' % n['nakladnik'])
        kolicine = {}
        for s in n['stavke']:
            b = obrasci.barkod_iz(s['barkod'])
            if b:
                kolicine[b] = kolicine.get(b, 0) + s['kolicina']
        sadrzaj, upisano, listovi = obrasci.popuni(o['sadrzaj'], kolicine)
        nema = [s for s in n['stavke'] if obrasci.barkod_iz(s['barkod']) not in upisano]
        ime = re.sub(r'(?i)\.xlsx$', '', o['ime']) + '_br%d-%d.xlsx' % (n['broj'], n['godina'])
        return {'ime': ime, 'sadrzaj': base64.b64encode(sadrzaj).decode(), 'upisano': len(upisano), 'nema': nema,
                'listovi': listovi}
    if put == '/api/narudzbenica/obrisi':
        c.execute('DELETE FROM narudzbenice_stavke WHERE narudzbenica_id = ?', (d['id'],))
        c.execute('DELETE FROM narudzbenice WHERE id = ?', (d['id'],))
        return {}
    if put == '/api/nakladnik':
        c.execute('INSERT OR REPLACE INTO nakladnici(naziv, email, kontakt) VALUES(?, ?, ?)',
                  (d['naziv'], d.get('email', ''), d.get('kontakt', '')))
        return {}
    if put == '/api/primka':
        return primka(c, d['stavke'])
    if put == '/api/rabljena':
        polja = ('knjiga_id', 'naziv', 'barkod', 'ocuvanost', 'otkup_cijena', 'prodajna_cijena', 'datum_otkupa',
                 'prodavatelj', 'telefon', 'status', 'datum_prodaje', 'kupac')
        if not (d.get('naziv') or '').strip():
            raise ValueError('Upiši naziv knjige.')
        d.setdefault('datum_otkupa', datetime.date.today().isoformat())
        d.setdefault('status', 'na stanju')
        if d.get('status') == 'prodano' and not d.get('datum_prodaje'):
            d['datum_prodaje'] = datetime.date.today().isoformat()
        vr = [d.get(p) for p in polja]
        if d.get('id'):
            c.execute('UPDATE rabljene SET ' + ', '.join(p + ' = ?' for p in polja) + ' WHERE id = ?', vr + [d['id']])
            return {'id': d['id']}
        return {'id': c.execute('INSERT INTO rabljene(' + ', '.join(polja) + ') VALUES(' + ', '.join('?' * len(polja)) + ')',
                                vr).lastrowid}
    if put == '/api/rabljena/obrisi':
        c.execute('DELETE FROM rabljene WHERE id = ?', (d['id'],))
        return {}
    raise KeyError(put)


def prijedlog_narudzbenice(c, nakladnik, razina=''):
    """Koliko treba naručiti: naručeno od kupaca (status 'naručeno', tekuća sezona) minus već naručeno od nakladnika."""
    god = sezona(c)
    trazi = rows(c.execute(
        'SELECT ' + KLJUC.format('st.') + ' AS kljuc, MAX(st.naziv) AS naziv, MAX(st.barkod) AS barkod, '
        "COALESCE(NULLIF(TRIM(MAX(st.nakladnik)), ''), 'bez nakladnika') AS nakladnik, MAX(st.cijena) AS cijena, "
        'SUM(st.kolicina) AS trazeno FROM stavke st JOIN narudzbe n ON n.id = st.narudzba_id '
        "WHERE n.godina = ? AND n.status = 'naručeno'" + (' AND n.skola_id IN (SELECT id FROM skole WHERE tip = ?)' if razina else '')
        + ' GROUP BY kljuc ORDER BY naziv', (god, razina) if razina else (god,)))
    # već naručeno: za OŠ/SŠ broje se samo narudžbenice te razine, za "sve" sve narudžbenice
    vec = {r['kljuc']: r['kolicina'] for r in c.execute(
        'SELECT ' + KLJUC.format('s.') + ' AS kljuc, SUM(s.kolicina) AS kolicina FROM narudzbenice_stavke s '
        "JOIN narudzbenice n ON n.id = s.narudzbenica_id WHERE n.godina = ? AND n.tip = 'narudzba'"
        + (' AND n.razina = ?' if razina else '') + ' GROUP BY kljuc', (god, razina) if razina else (god,))}
    stanje = {}
    try:
        stanje = {a['barkod']: a['stanje'] for a in vulpes_artikli(c) if a['barkod']}
    except ValueError:
        pass
    for r in trazi:
        r['vec_naruceno'] = vec.get(r['kljuc'], 0)
        r['stanje_vulpes'] = stanje.get((r['barkod'] or '').strip())
        r['kolicina'] = max(0, r['trazeno'] - r['vec_naruceno'])
    if nakladnik:
        return [r for r in trazi if r['nakladnik'] == nakladnik]
    sazetak = {}
    for r in trazi:
        s = sazetak.setdefault(r['nakladnik'], {'nakladnik': r['nakladnik'], 'naslova': 0, 'komada': 0})
        if r['kolicina']:
            s['naslova'] += 1
            s['komada'] += r['kolicina']
    return sorted(sazetak.values(), key=lambda s: (-s['komada'], s['nakladnik']))


def prijedlog_povrata(c, nakladnik):
    """Višak za povrat: naručeno od nakladnika − vraćeno − ono što trebaju kupci (sve osim otkazanih)."""
    god = sezona(c)
    nar = rows(c.execute(
        'SELECT ' + KLJUC.format('s.') + ' AS kljuc, MAX(s.naziv) AS naziv, MAX(s.barkod) AS barkod, MAX(s.cijena) AS cijena, '
        "SUM(CASE WHEN n.tip = 'narudzba' THEN s.kolicina ELSE -s.kolicina END) AS naruceno FROM narudzbenice_stavke s "
        'JOIN narudzbenice n ON n.id = s.narudzbenica_id WHERE n.godina = ? AND n.nakladnik = ? GROUP BY kljuc ORDER BY naziv',
        (god, nakladnik)))
    kupci = {r['kljuc']: r['kol'] for r in c.execute(
        'SELECT ' + KLJUC.format('st.') + ' AS kljuc, SUM(st.kolicina) AS kol FROM stavke st JOIN narudzbe n ON n.id = st.narudzba_id '
        "WHERE n.godina = ? AND n.status != 'otkazano' GROUP BY kljuc", (god,))}
    stanje = {}
    try:
        stanje = {a['barkod']: a['stanje'] for a in vulpes_artikli(c) if a['barkod']}
    except ValueError:
        pass
    out = []
    for r in nar:
        r['kupci'] = kupci.get(r['kljuc'], 0)
        r['stanje_vulpes'] = stanje.get((r['barkod'] or '').strip())
        r['kolicina'] = max(0, r['naruceno'] - r['kupci'])
        out.append(r)
    return out


def kupac_po_telefonu(c, telefon):
    znam = re.sub(r'\D', '', telefon or '')[-8:]
    if len(znam) < 6:
        return []
    return [r for r in rows(c.execute(
        'SELECT n.kupac, n.telefon, n.skola_id, s.naziv AS skola, n.razred, n.godina FROM narudzbe n '
        'LEFT JOIN skole s ON s.id = n.skola_id ORDER BY n.godina DESC, n.id DESC LIMIT 5000'))
        if re.sub(r'\D', '', r['telefon'] or '').endswith(znam)][:5]


def uvoz_mzo(c, ime, sadrzaj):
    """Ministarski katalog odobrenih udžbenika (xlsx): šifra kompleta, naziv, autori, nakladnik, razred, cijena."""
    redovi = popis.redovi_iz_datoteke(ime, sadrzaj)
    mapa, novi = None, []
    for red in redovi:
        z = popis.zaglavlje(red)
        if z:
            mapa = z
            continue
        if not mapa:
            continue
        g = lambda k: red[mapa[k]] if k in mapa and mapa[k] < len(red) else ''
        if not g('naziv'):
            continue
        try:
            cijena = float(str(g('cijena')).replace(',', '.').replace('€', '').strip() or 0)
        except ValueError:
            cijena = 0
        novi.append(((re.search(r'\d+', g('sifra')) or [''])[0], g('naziv'), g('autor'), g('nakladnik'), g('razred'),
                     g('predmet'), cijena))
    if not novi:
        raise ValueError('U datoteci nisam našao tablicu s nazivima udžbenika (stupac "Naziv").')
    c.execute('DELETE FROM mzo_katalog')
    c.executemany('INSERT INTO mzo_katalog VALUES(?, ?, ?, ?, ?, ?, ?)', novi)
    # poveži naše knjige sa šiframa iz kataloga po nazivu
    katalog = rows(c.execute("SELECT id, naziv, cijena, nakladnik, sifra_mzo FROM knjige WHERE COALESCE(sifra_mzo, '') = ''"))
    kand, _ = popis.kandidati(katalog, [])
    povezano = 0
    for sifra, naziv, autori, nakladnik, *_ in novi:
        if not sifra:
            continue
        p = popis.predlozi({'naslov': naziv, 'autori': autori, 'nakladnik': nakladnik, 'sifra': ''}, kand, {})
        isti_nakl = not (nakladnik and p and p[0].get('nakladnik')) or \
            popis.bez_kvacica(p[0]['nakladnik'])[:3] in popis.bez_kvacica(nakladnik)
        if p and p[0]['tip'] == 'k' and p[0]['ocjena'] >= 0.9 and isti_nakl:
            c.execute('UPDATE knjige SET sifra_mzo = ? WHERE id = ? AND COALESCE(sifra_mzo, \'\') = \'\'', (sifra, p[0]['id']))
            povezano += 1
    return {'udzbenika': len(novi), 'povezano': povezano}


def primka(c, stavke):
    """Raspodijeli pristigle knjige na narudžbe (najstarije prve); vrati narudžbe koje su sad kompletne."""
    god = sezona(c)
    rez = []
    for s in stavke:
        ostalo = int(s.get('kolicina') or 1)
        barkod = (s.get('barkod') or '').strip()
        redovi = c.execute(
            'SELECT st.id, st.kolicina - st.stiglo AS fali FROM stavke st JOIN narudzbe n ON n.id = st.narudzba_id '
            "WHERE n.godina = ? AND n.status = 'naručeno' AND st.stiglo < st.kolicina AND "
            '((? != \'\' AND st.barkod = ?) OR (st.knjiga_id IS NOT NULL AND st.knjiga_id = ?) OR st.naziv = ?) '
            'ORDER BY n.datum, n.id', (god, barkod, barkod, s.get('knjiga_id'), s.get('naziv', ''))).fetchall()
        rasp = 0
        for r in redovi:
            if not ostalo:
                break
            n = min(ostalo, r['fali'])
            c.execute('UPDATE stavke SET stiglo = stiglo + ? WHERE id = ?', (n, r['id']))
            ostalo -= n
            rasp += n
        rez.append({'naziv': s.get('naziv'), 'raspodijeljeno': rasp, 'visak': ostalo})
    kompletne = lista_narudzbi(c, god, 'naručeno', '')
    nepotpune = {r[0] for r in c.execute('SELECT DISTINCT narudzba_id FROM stavke WHERE stiglo < kolicina')}
    return {'stavke': rez, 'kompletne': [n for n in kompletne if n['id'] not in nepotpune]}


def statistika(c, god):
    w = (god,)
    uk = dict(c.execute(
        'SELECT COUNT(DISTINCT n.id) AS narudzbi, COALESCE(SUM(st.kolicina), 0) AS komada, '
        'COALESCE(SUM(st.kolicina * st.cijena), 0) AS iznos FROM narudzbe n LEFT JOIN stavke st ON st.narudzba_id = n.id '
        'WHERE n.godina = ?', w).fetchone())
    uk['uplaceno'] = c.execute('SELECT COALESCE(SUM(akontacija), 0) FROM narudzbe WHERE godina = ?', w).fetchone()[0]
    po = lambda sql: rows(c.execute(sql, w))
    return {
        'ukupno': uk,
        'po_statusu': po('SELECT status, COUNT(*) AS broj FROM narudzbe WHERE godina = ? GROUP BY status'),
        'po_platitelju': po("SELECT COALESCE(NULLIF(n.platitelj, ''), 'roditelj') AS platitelj, COUNT(DISTINCT n.id) AS narudzbi, "
                            'COALESCE(SUM(st.kolicina * st.cijena), 0) AS iznos FROM narudzbe n '
                            'LEFT JOIN stavke st ON st.narudzba_id = n.id WHERE n.godina = ? GROUP BY 1 ORDER BY iznos DESC'),
        'po_skoli': po("SELECT COALESCE(s.naziv, '—') AS skola, COUNT(DISTINCT n.id) AS narudzbi, "
                       'COALESCE(SUM(st.kolicina * st.cijena), 0) AS iznos FROM narudzbe n LEFT JOIN skole s ON s.id = n.skola_id '
                       'LEFT JOIN stavke st ON st.narudzba_id = n.id WHERE n.godina = ? GROUP BY s.naziv ORDER BY iznos DESC'),
        'po_nakladniku': po("SELECT COALESCE(NULLIF(st.nakladnik, ''), '—') AS nakladnik, SUM(st.kolicina) AS komada, "
                            'SUM(st.kolicina * st.cijena) AS iznos FROM stavke st JOIN narudzbe n ON n.id = st.narudzba_id '
                            'WHERE n.godina = ? GROUP BY 1 ORDER BY iznos DESC'),
        'top_knjige': po('SELECT st.naziv, st.nakladnik, SUM(st.kolicina) AS komada FROM stavke st '
                         'JOIN narudzbe n ON n.id = st.narudzba_id WHERE n.godina = ? GROUP BY st.naziv '
                         'ORDER BY komada DESC LIMIT 20'),
        'po_danu': po('SELECT datum, COUNT(*) AS broj FROM narudzbe WHERE godina = ? GROUP BY datum ORDER BY datum'),
        'rabljene': dict(c.execute(
            "SELECT COUNT(*) AS otkupljeno, COALESCE(SUM(otkup_cijena), 0) AS otkup, "
            "COALESCE(SUM(CASE WHEN status = 'prodano' THEN 1 END), 0) AS prodano, "
            "COALESCE(SUM(CASE WHEN status = 'prodano' THEN prodajna_cijena END), 0) AS prodaja "
            "FROM rabljene WHERE substr(datum_otkupa, 1, 4) = ?", (str(god),)).fetchone()),
    }


def spremi_narudzbu(c, d):
    stavke = [s for s in d.get('stavke', []) if (s.get('naziv') or '').strip()]
    if not (d.get('kupac') or '').strip():
        raise ValueError('Upiši ime i prezime kupca.')
    if not stavke:
        raise ValueError('Narudžba nema nijednu knjigu.')
    zaglavlje = (d['kupac'].strip(), d.get('telefon', ''), d.get('skola_id'), d.get('razred', ''),
                 float(d.get('akontacija') or 0), d.get('nacin_pl', ''), d.get('napomena', ''),
                 d.get('status') or STATUSI[0], (d.get('platitelj') or '').strip())
    if d.get('id'):
        nid = d['id']
        c.execute('UPDATE narudzbe SET kupac = ?, telefon = ?, skola_id = ?, razred = ?, akontacija = ?, '
                  'nacin_pl = ?, napomena = ?, status = ?, platitelj = ? WHERE id = ?', zaglavlje + (nid,))
        c.execute('DELETE FROM stavke WHERE narudzba_id = ?', (nid,))
    else:
        danas, god = datetime.date.today(), sezona(c)
        broj = c.execute('SELECT COALESCE(MAX(broj), 0) + 1 FROM narudzbe WHERE godina = ?', (god,)).fetchone()[0]
        nid = c.execute('INSERT INTO narudzbe(kupac, telefon, skola_id, razred, akontacija, nacin_pl, napomena, status, '
                        'platitelj, godina, broj, datum) VALUES(?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)',
                        zaglavlje + (god, broj, danas.isoformat())).lastrowid
    c.executemany('INSERT INTO stavke(narudzba_id, knjiga_id, naziv, barkod, nakladnik, cijena, kolicina, stiglo) '
                  'VALUES(?, ?, ?, ?, ?, ?, ?, ?)',
                  [(nid, s.get('knjiga_id'), s['naziv'], s.get('barkod', ''), s.get('nakladnik', ''),
                    float(s.get('cijena') or 0), int(s.get('kolicina') or 1),
                    min(int(s.get('stiglo') or 0), int(s.get('kolicina') or 1))) for s in stavke])
    if zaglavlje[7] == 'stiglo':
        c.execute('UPDATE narudzbe SET stiglo_datum = COALESCE(stiglo_datum, ?) WHERE id = ?',
                  (datetime.date.today().isoformat(), nid))
    return {'id': nid}


# ---------------------------------------------------------------- uvoz popisa knjiga po školama (Excel / CSV)

PREDLOZAK = ['Škola', 'Smjer', 'Razred', 'Barkod', 'Naziv', 'Nakladnik', 'Cijena']


def citaj_tablicu(ime, sadrzaj):
    if ime.lower().endswith('.xlsx'):
        return popis.citaj_xlsx(sadrzaj)
    for enc in ('utf-8-sig', 'cp1250'):
        try:
            tekst = sadrzaj.decode(enc)
            break
        except UnicodeDecodeError:
            pass
    sep = ';' if tekst.count(';') >= tekst.count(',') else ','
    return list(csv.reader(io.StringIO(tekst), delimiter=sep))


def uvoz_kompleta(c, ime, sadrzaj):
    redovi = [[str(x or '').strip() for x in r] for r in citaj_tablicu(ime, sadrzaj)]
    redovi = [r for r in redovi if any(r)]
    if not redovi:
        raise ValueError('Datoteka je prazna.')
    zaglavlje = [h.lower() for h in redovi[0]]
    def stupac(*imena):
        for i, h in enumerate(zaglavlje):
            if any(h.startswith(x) for x in imena):
                return i
        return None
    st = {k: stupac(*v) for k, v in {'skola': ('škola', 'skola'), 'smjer': ('smjer', 'program', 'komplet'),
                                       'razred': ('razred',), 'barkod': ('barkod', 'bar kod', 'isbn', 'ean'),
                                       'naziv': ('naziv', 'udžbenik', 'udzbenik'), 'nakladnik': ('nakladnik',),
                                       'cijena': ('cijena',)}.items()}
    if st['skola'] is None or st['razred'] is None or (st['naziv'] is None and st['barkod'] is None):
        raise ValueError('U prvom redu trebaju biti stupci: ' + ', '.join(PREDLOZAK) + '. Preuzmi predložak.')
    g = lambda r, k: r[st[k]] if st[k] is not None and st[k] < len(r) else ''
    vulpes = {}
    try:
        vulpes = {a['barkod']: a for a in vulpes_artikli(c) if a['barkod']}
    except ValueError:
        pass
    ociscen, novih_knjiga, kompleti, veza = set(), 0, set(), 0
    for r in redovi[1:]:
        sid = skola_id(c, g(r, 'skola'))
        smjer = g(r, 'smjer') or 'Komplet'
        razred = re.sub(r'\D', '', g(r, 'razred')) or g(r, 'razred')
        barkod = re.sub(r'[^0-9X]', '', g(r, 'barkod').upper())
        naziv = g(r, 'naziv')
        if not sid or not (naziv or barkod):
            continue
        k = c.execute('SELECT id FROM kompleti WHERE skola_id = ? AND naziv = ? AND razred = ?', (sid, smjer, razred)).fetchone()
        kid = k['id'] if k else c.execute('INSERT INTO kompleti(naziv, skola_id, razred) VALUES(?, ?, ?)',
                                          (smjer, sid, razred)).lastrowid
        if kid not in ociscen:  # novi popis zamjenjuje stari za taj komplet
            novi_komplet_popis(c, kid)
            ociscen.add(kid)
        kompleti.add(kid)
        knj = None
        if barkod:
            knj = c.execute('SELECT id FROM knjige WHERE barkod = ?', (barkod,)).fetchone()
        if not knj and naziv:
            knj = c.execute('SELECT id FROM knjige WHERE UPPER(naziv) = UPPER(?)', (naziv,)).fetchone()
        if knj:
            knjiga = knj['id']
            c.execute('UPDATE knjige SET aktivna = 1 WHERE id = ?', (knjiga,))
        else:
            v = vulpes.get(barkod) or {}
            try:
                cijena = float(g(r, 'cijena').replace(',', '.') or 0)
            except ValueError:
                cijena = 0
            knjiga = c.execute('INSERT INTO knjige(naziv, nakladnik, cijena, barkod, vulpes_sifra) VALUES(?, ?, ?, ?, ?)',
                               (naziv or v.get('naziv', barkod), g(r, 'nakladnik'), v.get('cijena') or cijena,
                                barkod, v.get('sifra'))).lastrowid
            novih_knjiga += 1
        c.execute('INSERT OR IGNORE INTO komplet_knjige(komplet_id, knjiga_id) VALUES(?, ?)', (kid, knjiga))
        veza += 1
    return {'kompleta': len(kompleti), 'knjiga': veza, 'novih_knjiga': novih_knjiga}


def spremi_popis(c, d):
    sid = d.get('skola_id')
    if not sid:
        raise ValueError('Odaberi školu.')
    br_kompleta, br_novih = 0, 0
    for sek in d['sekcije']:
        stavke = [s for s in sek['stavke'] if s.get('izbor')]
        if not stavke:
            continue
        smjer = (sek.get('smjer') or '').strip() or 'Komplet'
        razred = (sek.get('razred') or '').strip()
        r = c.execute('SELECT id FROM kompleti WHERE skola_id = ? AND naziv = ? AND razred = ?', (sid, smjer, razred)).fetchone()
        kid = r['id'] if r else c.execute('INSERT INTO kompleti(naziv, skola_id, razred) VALUES(?, ?, ?)',
                                          (smjer, sid, razred)).lastrowid
        novi_komplet_popis(c, kid)
        br_kompleta += 1
        for s in stavke:
            iz = s['izbor']
            if iz['tip'] == 'k':
                knjiga = iz['id']
            elif iz['tip'] == 'v':
                r = c.execute('SELECT id FROM knjige WHERE barkod = ? AND barkod != ""', (iz['barkod'],)).fetchone()
                knjiga = r['id'] if r else c.execute(
                    'INSERT INTO knjige(naziv, cijena, barkod, vulpes_sifra) VALUES(?, ?, ?, ?)',
                    (iz['naziv'], iz['cijena'], iz['barkod'], iz['sifra'])).lastrowid
                br_novih += 0 if r else 1
            else:
                knjiga = c.execute('INSERT INTO knjige(naziv, autori, nakladnik) VALUES(?, ?, ?)',
                                   (iz.get('naziv') or s.get('naslov'), s.get('autori', ''), s.get('nakladnik', ''))).lastrowid
                br_novih += 1
            c.execute('UPDATE knjige SET aktivna = 1 WHERE id = ?', (knjiga,))
            if s.get('sifra'):
                c.execute('UPDATE knjige SET sifra_mzo = ? WHERE id = ?', (s['sifra'], knjiga))
            c.execute('INSERT OR IGNORE INTO komplet_knjige(komplet_id, knjiga_id, predmet) VALUES(?, ?, ?)',
                      (kid, knjiga, (s.get('predmet') or '').strip()[:60]))
    return {'kompleta': br_kompleta, 'novih_knjiga': br_novih}


def csv_odgovor(naslovi, redovi):
    buf = io.StringIO()
    w = csv.writer(buf, delimiter=';')
    w.writerow(naslovi)
    for r in redovi:
        w.writerow([('%.2f' % v).replace('.', ',') if isinstance(v, float) else v for v in r])
    return ('\ufeff' + buf.getvalue()).encode('utf-8')


def api_csv(c, put, q):
    g = lambda k: q.get(k, [''])[0]
    if put == '/csv/narudzbe':
        ns = lista_narudzbi(c, g('godina'), g('status'), g('skola'), g('platitelj', '*'), g('razina'))
        return csv_odgovor(['Broj', 'Datum', 'Kupac', 'Telefon', 'Škola', 'Razred', 'Komada', 'Iznos', 'Akontacija',
                            'Ostatak', 'Status'],
                           [[n['broj'], n['datum'], n['kupac'], n['telefon'], n['skola'], n['razred'], n['komada'],
                             float(n['iznos']), float(n['akontacija'] or 0),
                             float(n['iznos'] - (n['akontacija'] or 0)), n['status']] for n in ns])
    if put == '/csv/narudzbenica':
        n = api_get(c, '/api/narudzbenica', q)
        return csv_odgovor(['Barkod', 'Naziv', 'Količina'], [[s['barkod'], s['naziv'], s['kolicina']] for s in n['stavke']])
    if put == '/csv/predlozak':
        return csv_odgovor(PREDLOZAK, [['SŠ METKOVIĆ', 'Komercijalist', '1', '9789530226043', 'OSNOVE STROJARSTVA', 'ŠK', '28,00']])
    if put == '/csv/platitelj':
        ns = api_get(c, '/api/platitelj', q)
        return csv_odgovor(['Br. narudžbe', 'Učenik / kupac', 'Telefon', 'Škola', 'Razred', 'Knjiga', 'Nakladnik', 'Kol.', 'Cijena', 'Iznos'],
                           [[n['broj'], n['kupac'], n['telefon'], n['skola'], n['razred'], s['naziv'], s['nakladnik'], s['kolicina'],
                             s['cijena'], s['cijena'] * s['kolicina']] for n in ns for s in n['stavke']])
    if put == '/csv/zbirno':
        zs = zbirno(c, g('godina'), g('status'), g('skola'), g('razina'))
        return csv_odgovor(['Nakladnik', 'Naziv', 'Barkod', 'Količina', 'Cijena'],
                           [[z['nakladnik'], z['naziv'], z['barkod'], z['kolicina'], float(z['cijena'])] for z in zs])
    raise KeyError(put)


class Handler(BaseHTTPRequestHandler):
    def log_message(self, *a):
        pass

    def posalji(self, kod, tijelo, tip='application/json; charset=utf-8', ime=None):
        self.send_response(kod)
        self.send_header('Content-Type', tip)
        self.send_header('Cache-Control', 'no-store')
        if ime:
            self.send_header('Content-Disposition', 'attachment; filename="%s"' % ime)
        self.send_header('Content-Length', str(len(tijelo)))
        self.end_headers()
        self.wfile.write(tijelo)

    def obradi(self, fn):
        try:
            with transakcija() as c:
                rez = fn(c)
            if isinstance(rez, bytes):
                self.posalji(200, rez, 'text/csv; charset=utf-8', urlparse(self.path).path.split('/')[-1] + '.csv')
            else:
                self.posalji(200, json.dumps(rez, ensure_ascii=False, default=str).encode('utf-8'))
        except KeyError:
            self.posalji(404, b'{"greska": "Nepoznata adresa"}')
        except Exception as e:  # poruka ide korisniku u preglednik
            self.posalji(400, json.dumps({'greska': str(e)}, ensure_ascii=False).encode('utf-8'))

    def do_GET(self):
        u = urlparse(self.path)
        if u.path in ('/', '/index.html'):
            with open(os.path.join(MAPA, 'index.html'), 'rb') as f:
                return self.posalji(200, f.read(), 'text/html; charset=utf-8')
        if u.path == '/logo.png':  # logo na ispisima - zamijenite datoteku logo.png u mapi programa
            put = os.path.join(MAPA, 'logo.png')
            if os.path.exists(put):
                with open(put, 'rb') as f:
                    return self.posalji(200, f.read(), 'image/png')
            return self.posalji(404, b'')
        q = parse_qs(u.query)
        if u.path.startswith('/csv/'):
            return self.obradi(lambda c: api_csv(c, u.path, q))
        self.obradi(lambda c: api_get(c, u.path, q))

    def do_POST(self):
        u = urlparse(self.path)
        d = json.loads(self.rfile.read(int(self.headers.get('Content-Length') or 0)) or b'{}')
        self.obradi(lambda c: api_post(c, u.path, d))


# ---------------------------------------------------------------- prijenos iz starog programa

def uvoz_starog(mapa):
    def tablica(ime):
        for f in os.listdir(mapa):
            if f.lower() == ime.lower() + '.dbf':
                return dbf.citaj(os.path.join(mapa, f))
        return []

    with transakcija() as c:
        if c.execute('SELECT COUNT(*) FROM knjige').fetchone()[0]:
            raise SystemExit('Baza već ima knjige; prijenos se radi samo u praznu bazu (obriši knjige.db).')
        po_broju = {}
        for a in tablica('artikli'):
            po_broju[a['BROJ']] = c.execute(
                'INSERT INTO knjige(naziv, vrsta, autori, nakladnik, cijena, barkod, aktivna, stari_broj) '
                'VALUES(?, ?, ?, ?, ?, ?, ?, ?)',
                (a['NAZIV'], a['OZNAKA'], a['AUTORI'], a['NAKLADNIK'], a['CIJENA'] or 0, a['BARKOD'],
                 1 if a['AKTIVNA'] else 0, a['BROJ'])).lastrowid
        kompleti = {}
        for k in tablica('komplet'):
            kompleti[k['BROJ']] = c.execute('INSERT INTO kompleti(naziv, skola_id, razred) VALUES(?, ?, ?)',
                                            (k['NAZIV'], skola_id(c, k['SKOLA']), k['RAZRED'])).lastrowid
        for s in tablica('komplets'):
            if s['BROJ'] in kompleti and s['BR_ROBE'] in po_broju:
                c.execute('INSERT OR IGNORE INTO komplet_knjige(komplet_id, knjiga_id) VALUES(?, ?)',
                          (kompleti[s['BROJ']], po_broju[s['BR_ROBE']]))
        narudzbe = {}
        for n in tablica('nar'):
            datum = n['DATUM'] or datetime.date.today()
            kupac = ' '.join(x for x in (n['IME'], n['PREZIME']) if x)
            narudzbe[n['BROJ']] = c.execute(
                'INSERT INTO narudzbe(godina, broj, datum, kupac, telefon, skola_id, razred, akontacija, nacin_pl, '
                'napomena, status) VALUES(?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)',
                (datum.year, n['BR_NAR'] or n['BROJ'], datum.isoformat(), kupac, n['TELEFON'],
                 skola_id(c, n['SKOLA']), n['RAZRED'].strip(' ,'), n['UPLACENO'] or 0, n['NACIN_PL'],
                 n.get('NAPOMENA') or '', 'preuzeto')).lastrowid
        stavke = {}
        for s in tablica('nars'):
            if s['BROJ'] not in narudzbe:
                continue
            kljuc = (s['BROJ'], s['BR_ROBE'], s['NAZIV'])
            if kljuc in stavke:
                c.execute('UPDATE stavke SET kolicina = kolicina + 1 WHERE id = ?', (stavke[kljuc],))
            else:
                stavke[kljuc] = c.execute(
                    'INSERT INTO stavke(narudzba_id, knjiga_id, naziv, barkod, nakladnik, cijena, kolicina) '
                    'VALUES(?, ?, ?, ?, ?, ?, 1)',
                    (narudzbe[s['BROJ']], po_broju.get(s['BR_ROBE']), s['NAZIV'], s['BARKOD'], s['NAKLADNIK'],
                     s['CIJENA'] or 0)).lastrowid
        for r in tablica('napomena')[:1]:
            c.execute("UPDATE postavke SET v = ? WHERE k = 'napomena_listic'", (r.get('NAPOMENA') or '',))
        print('Preneseno: %d knjiga, %d kompleta, %d narudžbi, %d stavki.'
              % (len(po_broju), len(kompleti), len(narudzbe), len(stavke)))


def main():
    if sys.stdout is None:  # pokrenuto bez prozora (pythonw) - poruke idu u datoteku
        sys.stdout = sys.stderr = open(os.path.join(MAPA, 'program.log'), 'a', encoding='utf-8', buffering=1)
    init_db()
    if '--provjera' in sys.argv:  # nadogradnja provjerava može li se nova verzija pokrenuti
        print('ok', VERZIJA)
        return
    if len(sys.argv) > 2 and sys.argv[1] == '--uvoz':
        return uvoz_starog(sys.argv[2])
    cfg = {'port': 8765, 'mreza': False, 'javni_port': 8766}
    if os.path.exists(CONFIG):
        with open(CONFIG, encoding='utf-8') as f:
            cfg.update(json.load(f))
    else:
        with open(CONFIG, 'w', encoding='utf-8') as f:
            json.dump(cfg, f, indent=2)
    adresa = '0.0.0.0' if cfg['mreza'] else '127.0.0.1'
    url = 'http://localhost:%d/' % cfg['port']
    import socket
    if '--nakon-nadogradnje' in sys.argv:  # pričekaj da se stara verzija ugasi
        for _ in range(60):
            zauzeto = False
            for port in (int(cfg['port']), int(cfg.get('javni_port') or 0)):
                if port:
                    with socket.socket() as s:
                        zauzeto = zauzeto or s.connect_ex(('127.0.0.1', port)) == 0
            if not zauzeto:
                break
            time.sleep(0.5)
    with socket.socket() as s:  # program već radi - samo otvori preglednik
        if s.connect_ex(('127.0.0.1', int(cfg['port']))) == 0:
            if '--bez-preglednika' not in sys.argv:
                webbrowser.open(url)
            return
    srv = ThreadingHTTPServer((adresa, int(cfg['port'])), Handler)
    Handler.server_ref = srv
    if cfg.get('javni_port'):  # online predbilježba (samo taj dio se objavljuje na internetu)
        try:
            javno.pokreni(transakcija, int(cfg['javni_port']))
            print('Predbilježba za roditelje radi na http://localhost:%d/' % int(cfg['javni_port']))
        except OSError as e:
            print('Predbilježba se nije pokrenula:', e)
    print('Program radi na', url, '- ne zatvarajte ovaj prozor.')
    if '--bez-preglednika' not in sys.argv:
        threading.Thread(target=lambda: (time.sleep(1), webbrowser.open(url)), daemon=True).start()
    threading.Thread(target=kopije_u_pozadini, daemon=True).start()  # jedna automatska kopija dnevno
    srv.serve_forever()


if __name__ == '__main__':
    main()

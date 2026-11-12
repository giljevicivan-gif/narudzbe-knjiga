"""Čitanje popisa udžbenika koje objavljuju škole (PDF, Word, Excel, CSV) i prijedlog kompleta.

Svaka datoteka se pretvori u listu redova, a red je lista ćelija (odlomak teksta = red s jednom ćelijom).
"""
import csv
import io
import os
import re
import sys
import unicodedata
import xml.etree.ElementTree as ET
import zipfile

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), 'pypdf.zip'))

NAKLADNICI = ['ALFA', 'PROFIL', 'KLETT', 'ŠKOLSKA KNJIGA', 'ŠK', 'ELEMENT', 'NEODIDACTA', 'LJEVAK', 'SALESIANA',
              'OXFORD', 'PEARSON', 'CAMBRIDGE', 'MACMILLAN', 'MERIDIJANI', 'SYSPRINT', 'ZNAMEN', 'KRŠĆANSKA SADAŠNJOST',
              'HRVATSKA SVEUČILIŠNA', 'EDUKA', 'ALKA', 'GLAS KONCILA', 'MOZAIK', 'HUEBER', 'ZAVOD ZA UDŽBENIKE',
              'ARCA', 'DIDAKTA', 'TEHNIČKA KNJIGA', 'PLANET', 'EXPRESS', 'NAKLADA', 'Š. K.', 'Š.K.', 'S.K.']
NAKLADNIK_RE = re.compile(r'(?<!\w)(' + '|'.join(re.escape(n) for n in sorted(set(NAKLADNICI), key=len, reverse=True)) + r')(?!\w)')
RAZRED_RE = re.compile(r'\b([1-5])\s*\.\s*(?:[a-f]\s+)?razred|\brazred\s*:?\s*([1-5])\b|\b(I{1,3}|IV)\s*\.\s*razred', re.I)
RIMSKI = {'I': '1', 'II': '2', 'III': '3', 'IV': '4'}
RADNA_RE = re.compile(r'RADN|R\.\s?B\.|-R\.?B|\bRB\b|WORKBOOK|ZBIRK|Z\.Z\.|VJEZBENIC|ARBEITSBUCH|QUADERNO|CAHIER')


# ------------------------------------------------------------- čitanje datoteka

def redovi_iz_datoteke(ime, sadrzaj):
    ime = ime.lower()
    if ime.endswith('.pdf'):
        redovi = _pdf(sadrzaj)
    elif ime.endswith('.docx'):
        redovi = _docx(sadrzaj)
    elif ime.endswith('.xlsx'):
        redovi = citaj_xlsx(sadrzaj)
    elif ime.endswith(('.doc', '.xls')):
        raise ValueError('Stari .doc / .xls format: otvori datoteku u Wordu/Excelu i spremi kao .docx / .xlsx ili PDF.')
    else:
        tekst = None
        for enc in ('utf-8-sig', 'cp1250'):
            try:
                tekst = sadrzaj.decode(enc)
                break
            except UnicodeDecodeError:
                pass
        sep = ';' if tekst.count(';') >= tekst.count('\t') else '\t'
        redovi = list(csv.reader(io.StringIO(tekst), delimiter=sep))
    redovi = [[re.sub(r'\s+', ' ', str(c or '')).strip() for c in r] for r in redovi]
    return [[c for c in r] for r in redovi if any(r)]


def citaj_xlsx(sadrzaj):
    z = zipfile.ZipFile(io.BytesIO(sadrzaj))
    m = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
    zajednicki = []
    if 'xl/sharedStrings.xml' in z.namelist():
        for si in ET.fromstring(z.read('xl/sharedStrings.xml')).findall(m + 'si'):
            zajednicki.append(''.join(t.text or '' for t in si.iter(m + 't')))
    listovi = sorted((n for n in z.namelist() if re.match(r'xl/worksheets/sheet\d+\.xml$', n)),
                     key=lambda n: int(re.search(r'(\d+)\.xml$', n).group(1)))
    redovi = []
    for list_ in listovi:
        for row in ET.fromstring(z.read(list_)).iter(m + 'row'):
            red = {}
            for cel in row.findall(m + 'c'):
                idx = 0
                for ch in re.match(r'[A-Z]+', cel.get('r')).group():
                    idx = idx * 26 + ord(ch) - 64
                v = cel.find(m + 'v')
                if cel.get('t') == 's' and v is not None:
                    val = zajednicki[int(v.text)]
                elif cel.get('t') == 'inlineStr':
                    val = ''.join(t.text or '' for t in cel.iter(m + 't'))
                else:
                    val = v.text if v is not None else ''
                    if re.fullmatch(r'\d+\.0', val or ''):
                        val = val[:-2]
                red[idx - 1] = val
            if red:
                redovi.append([red.get(i, '') for i in range(max(red) + 1)])
    return redovi


def _docx(sadrzaj):
    w = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
    tijelo = ET.fromstring(zipfile.ZipFile(io.BytesIO(sadrzaj)).read('word/document.xml')).find(w + 'body')
    redovi = []

    def tekst(el):
        return ' '.join(''.join(t.text or '' for t in p.iter(w + 't')) for p in el.iter(w + 'p')).strip()

    def obidi(el):
        for d in el:
            if d.tag == w + 'p':
                redovi.append([tekst(d)])
            elif d.tag == w + 'tbl':
                for tr in d.findall(w + 'tr'):
                    redovi.append([tekst(tc) for tc in tr.findall(w + 'tc')])
            elif d.tag in (w + 'sdt', w + 'sdtContent', w + 'customXml'):
                obidi(d)
    obidi(tijelo)
    return redovi


def _pdf(sadrzaj):
    """Tekst iz PDF-a s položajima; tablice se slažu po stupcima iz zaglavlja (red s 'NAZIV')."""
    from pypdf import PdfReader
    linije = []
    for n, stranica in enumerate(PdfReader(io.BytesIO(sadrzaj)).pages):
        dijelovi = []

        def posjet(t, cm, tm, font, vel):
            if t and t.strip():
                x = cm[0] * tm[4] + cm[2] * tm[5] + cm[4]
                y = cm[1] * tm[4] + cm[3] * tm[5] + cm[5]
                dijelovi.append([x, y, t.strip()])
        try:
            stranica.extract_text(visitor_text=posjet)
        except Exception:
            continue
        dijelovi.sort(key=lambda d: (-d[1], d[0]))
        for d in dijelovi:
            if linije and linije[-1]['str'] == n and abs(linije[-1]['y'] - d[1]) < 3:
                linije[-1]['d'].append(d)
            else:
                linije.append({'str': n, 'y': d[1], 'd': [d]})
    if not linije:
        raise ValueError('PDF nema teksta (vjerojatno je skeniran). Treba popis kao tekstualni PDF, Word ili Excel.')
    for l in linije:
        l['d'].sort(key=lambda d: d[0])
        spojeni = []  # spoji dijelove iste riječi/ćelije koji su blizu
        for d in l['d']:
            if spojeni and d[0] - spojeni[-1][3] < 4:
                spojeni[-1][2] += ('' if spojeni[-1][2].endswith('-') else ' ') + d[2]
                spojeni[-1][3] = d[0] + len(d[2]) * 4.5
            else:
                spojeni.append([d[0], d[1], d[2], d[0] + len(d[2]) * 4.5])
        l['d'] = spojeni

    redovi, i = [], 0
    while i < len(linije):
        l = linije[i]
        tekstovi = [d[2] for d in l['d']]
        if zaglavlje(tekstovi):
            i = _pdf_tablica(linije, i, redovi)
            continue
        redovi.append(tekstovi if len(tekstovi) > 1 else [' '.join(tekstovi)])
        i += 1
    return redovi


def _pdf_tablica(linije, i, redovi):
    """Od linije zaglavlja (indeks i) složi redove tablice; vraća indeks prve linije iza tablice."""
    zag = linije[i]['d']
    stupci = [d[0] for d in zag]
    nazivi = [d[2] for d in zag]
    sifra_st = next((j for j, t in enumerate(nazivi) if re.search(r'SIFRA|KAT|REG|\bBR\b', bez_kvacica(t))), None)
    redovi.append(nazivi)
    tablica = []
    i += 1
    while i < len(linije):
        l = linije[i]
        prvi_x = l['d'][0][0]
        cijela = ' '.join(d[2] for d in l['d'])
        if prvi_x < stupci[0] - 8:
            break
        if len(l['d']) == 1 and prvi_x <= stupci[0] + 10 and len(cijela) > 12 and \
                (RAZRED_RE.search(cijela) or re.search(r'(?i)ravnatelj|potpis|napomena|klasa', cijela)):
            break
        if zaglavlje([d[2] for d in l['d']]):
            break
        stanice = [''] * len(stupci)
        for d in l['d']:
            j = max([k for k, s in enumerate(stupci) if s <= d[0] + 6] or [0])
            stanice[j] = (stanice[j] + ' ' + d[2]).strip()
        pune = [s for s in stanice if s]
        if not tablica and len(pune) < len(stupci) - 1 and all(len(s.split()) <= 2 and not re.search(r'\d', s) for s in pune):
            i += 1  # drugi red zaglavlja (npr. "KOMPLETA", "UDŽBENIKA")
            continue
        tablica.append(stanice)
        i += 1
    if not tablica:
        return i
    puni = [sum(1 for s in st if s) >= max(2, len(stupci) - 2) for st in tablica]
    sa_sifrom = [bool(sifra_st is not None and re.search(r'\d{3,5}', st[sifra_st])) for st in tablica]
    pocetak = sa_sifrom if sum(sa_sifrom) >= max(1, sum(puni) * 0.5) else puni
    zadnji = None
    for st, novi in zip(tablica, pocetak):
        if novi or zadnji is None:
            zadnji = st
            redovi.append(zadnji)
        else:
            for j, s in enumerate(st):
                if s:
                    zadnji[j] = (zadnji[j] + ('' if zadnji[j].endswith('-') else ' ') + s).strip()
    return i


# ------------------------------------------------------------- prepoznavanje

def bez_kvacica(s):
    s = s.upper().replace('Đ', 'DJ')
    return ''.join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn')


RIJECI_VAN = {'ZA', 'I', 'U', 'IZ', 'S', 'SA', 'NA', 'OD', 'DO', 'UDZBENIK', 'UDZB', 'UDZ', 'RAZRED', 'RAZ', 'R',
              'SREDNJE', 'SREDNJIH', 'SKOLE', 'SKOLA', 'SK', 'GIMNAZIJE', 'GIM', 'GODINU', 'DIO', 'RADNA', 'BILJEZNICA',
              'BILJ', 'RB', 'SADRZAJEM', 'MULTIMEDIJSKIM', 'PRVOM', 'DRUGOM', 'TRECEM', 'CETVRTOM', 'PRVI', 'DRUGI',
              'TRECI', 'CETVRTI', 'STRUKOVNIH', 'STRUKOVNE', 'CETVEROGODISNJIH', 'TROGODISNJIH', 'THE', 'AND', 'WITH',
              'INTEGRIRANI', 'GODISNJE', 'SATI', 'PROFIL', 'ALFA', 'ELEMENT', 'KLETT', 'NEODIDACTA'}


def rijeci(s):
    return [r for r in re.findall(r'[A-Z]+|\d+', bez_kvacica(s)) if r not in RIJECI_VAN and (len(r) > 1 or r.isdigit())]


def _ima(r, skup, popis_):
    return r in skup or (len(r) >= 4 and any(x.startswith(r) or (len(x) >= 4 and r.startswith(x)) for x in popis_))


def ocjena(red, kand):
    """red/kand: dict s 'sve' (riječi), 'skup', 'naslov' (riječi glavnog naslova), 'mali' (brojevi), 'rb'."""
    rk = kand['sve']
    pokriveno = sum(1 for r in rk if _ima(r, red['skup'], red['sve'])) / len(rk)
    if red['naslov']:
        preciz = sum(1 for r in red['naslov'] if _ima(r, kand['skup'], rk)) / len(red['naslov'])
    else:
        preciz = pokriveno
    o = 0.5 * pokriveno + 0.5 * preciz
    if kand['mali'] - red['mali_sve'] or red['mali'] - kand['mali']:
        o *= 0.4
    if kand['rb'] != red['rb']:
        o -= 0.25
    return o


def _opis(naziv, naslov=None):
    sve = rijeci(naziv)
    glavni = re.split(r'\s[-–:]\s|:|\s-\s?|\(', naslov if naslov is not None else naziv)[0]
    n = rijeci(glavni)
    return {'sve': sve, 'skup': set(sve), 'naslov': n, 'mali': {r for r in n if r.isdigit() and len(r) <= 2},
            'mali_sve': {r for r in sve if r.isdigit() and len(r) <= 2}, 'rb': bool(RADNA_RE.search(bez_kvacica(naziv)))}


def razred_iz(tekst):
    m = RAZRED_RE.search(tekst or '')
    if not m:
        return ''
    return m.group(1) or m.group(2) or RIMSKI.get((m.group(3) or '').upper(), '')


def ocisti_smjer(tekst):
    t = re.sub(r'(?i)popis\s+(obveznih\s+)?(udžbenika|udzbenika)(\s+i\s+[a-zčćšžđ ]+)?|za\s+školsku\s+godinu|'
               r'šk(olsk[a-z]*)?\.?\s*god(in[a-z]*)?\.?|\d{4}\s*\.?\s*/\s*\d{2,4}\s*\.?|srednje\s+škole', ' ', tekst)
    t = RAZRED_RE.sub(' ', t)
    t = re.sub(r'(?i)\b(za|u)\b\s*$', ' ', t.strip())
    t = re.sub(r'[|:\-–]+', ' ', t)
    return re.sub(r'\s+', ' ', t).strip(' .,/')


KLJUCNI = {'naziv': ('NAZIV', 'UDZBENIK', 'NASLOV'), 'sifra': ('SIFRA', 'KAT', 'REG', 'BR.', 'BROJ'),
           'autor': ('AUTOR',), 'nakladnik': ('NAKLADNIK', 'IZDAVAC'), 'razred': ('RAZRED',),
           'predmet': ('PREDMET',), 'cijena': ('CIJENA',), 'smjer': ('SMJER', 'PROGRAM', 'ZANIMANJE')}


def zaglavlje(red):
    b = [bez_kvacica(c) for c in red]
    if len(red) < 3 or any(len(c) > 40 or re.search(r'\d', c) for c in b) \
            or not any('NAZIV' in c or c in ('UDZBENIK', 'NASLOV') for c in b) \
            or not any(re.search(r'AUTOR|NAKLADNIK|PREDMET|SIFRA|KAT|IZDAVAC', c) for c in b):
        return None
    mapa = {}
    for i, c in enumerate(red):
        b = bez_kvacica(c)
        for k, kljucevi in KLJUCNI.items():
            if k not in mapa and any(b.startswith(x) or (x in b and len(x) > 3) for x in kljucevi):
                mapa[k] = i
                break
    return mapa if 'naziv' in mapa else None


def analiziraj(redovi, katalog, vulpes):
    """Vraća sekcije [{smjer, razred, stavke: [{naslov, autori, nakladnik, sifra, prijedlozi}]}]."""
    sekcije, trenutna, mapa, zadnji_tekst, smjer, predmet = [], None, None, '', '', ''

    def sekcija(razred):
        nonlocal trenutna
        trenutna = {'smjer': smjer, 'razred': razred, 'stavke': []}
        sekcije.append(trenutna)

    for red in redovi:
        stanice = [c for c in red if c]
        if not stanice:
            continue
        z = zaglavlje(red)
        if z:
            mapa = z
            continue
        cijela = ' '.join(stanice)
        if len(stanice) == 1:
            raz = razred_iz(cijela)
            if re.fullmatch(r'[A-ZČĆŽŠĐ][A-ZČĆŽŠĐ ,.\-/]{2,40}( \d)?', cijela) and not raz \
                    and not re.search(r'ŠKOLA|SKOLA|GIMNAZIJ|RAZRED|NAKLADNIK|AUTOR', cijela):
                predmet = cijela.strip()  # naslov predmeta iznad knjiga (npr. "HRVATSKI JEZIK")
                continue
            if raz and len(cijela) < 160:
                predmet = ''
                smjer = ocisti_smjer(cijela) or ocisti_smjer(zadnji_tekst) or smjer
                sekcija(raz)
            elif mapa is None or len(cijela) > 3:
                zadnji_tekst = cijela
                if re.search(r'(?i)gimnazij|tehničar|tehnicar|ekonomist|komercijalist|smjer|program|zanimanj', cijela) \
                        and len(cijela) < 100 and not NAKLADNIK_RE.search(cijela.upper()):
                    smjer = ocisti_smjer(cijela)
                    if trenutna and not trenutna['stavke']:
                        trenutna['smjer'] = smjer
            continue
        g = lambda k: red[mapa[k]] if mapa and k in mapa and mapa[k] < len(red) else ''
        if mapa:
            naslov, autori, nakladnik, sifra = g('naziv'), g('autor'), g('nakladnik'), g('sifra')
            if not naslov or len(naslov) < 3:
                continue
        else:  # bez zaglavlja: red je knjiga ako ima nakladnika
            if not NAKLADNIK_RE.search(cijela.upper()):
                zadnji_tekst = cijela
                continue
            kand = [c for c in stanice if not re.fullmatch(r'[\d.,\s€kn]+', c) and not NAKLADNIK_RE.fullmatch(c.upper())]
            naslov = max(kand, key=len) if kand else cijela
            autori, nakladnik = '', next((c for c in stanice if NAKLADNIK_RE.search(c.upper())), '')
            sifra = next((c for c in stanice if re.fullmatch(r'\d{3,5}', c)), '')
        sifra = (re.search(r'\d{3,5}', sifra or '') or [''])[0] if sifra else ''
        raz = razred_iz(g('razred')) if mapa else ''
        if trenutna is None or (raz and trenutna['razred'] and raz != trenutna['razred'] and not trenutna['stavke']):
            sekcija(raz or razred_iz(zadnji_tekst))
        elif raz and not trenutna['razred']:
            trenutna['razred'] = raz
        trenutna['stavke'].append({'naslov': naslov[:200], 'autori': autori[:120], 'nakladnik': nakladnik[:60],
                                   'sifra': sifra, 'predmet': ((g('predmet') if mapa else '') or predmet)[:60]})

    sekcije = [s for s in sekcije if s['stavke']]
    kand, po_sifri = kandidati(katalog, vulpes)
    for s in sekcije:
        for st in s['stavke']:
            st['prijedlozi'] = predlozi(st, kand, po_sifri)
    return sekcije


def kandidati(katalog, vulpes):
    kand = [(_opis(k['naziv']), 0.05, {'tip': 'k', 'id': k['id'], 'naziv': k['naziv'], 'cijena': k['cijena'],
                                       'nakladnik': k.get('nakladnik') or ''}) for k in katalog]
    kand += [(_opis(v['naziv']), 0, {'tip': 'v', 'sifra': v['sifra'], 'barkod': v['barkod'], 'naziv': v['naziv'],
                                     'cijena': v['cijena']}) for v in vulpes if v['barkod'][:3] in ('978', '979')]
    return [k for k in kand if k[0]['sve']], {k['sifra_mzo']: k for k in katalog if k.get('sifra_mzo')}


def predlozi(st, kand, po_sifri):
    k = po_sifri.get(st['sifra'])
    if k:
        return [{'tip': 'k', 'id': k['id'], 'naziv': k['naziv'], 'cijena': k['cijena'], 'ocjena': 1.0}]
    red = _opis(st['naslov'] + ' ' + st['autori'], st['naslov'])
    if not red['sve']:
        return []
    prefiksi = {x[:4] for x in red['sve']}
    ocjene = []
    for opis, bonus, p in kand:
        if not (opis['skup'] & red['skup']) and not any(x[:4] in prefiksi for x in opis['sve'] if len(x) >= 4):
            continue
        o = ocjena(red, opis)
        if p.get('nakladnik') and st['nakladnik'] and bez_kvacica(p['nakladnik'])[:3] in bez_kvacica(st['nakladnik']):
            o += 0.05
        if o >= 0.45:
            ocjene.append((o + bonus, p))
    ocjene.sort(key=lambda x: -x[0])
    rez, vidjeno = [], set()
    for o, p in ocjene:
        kljuc = (p['tip'], p.get('id') or p.get('sifra'))
        if kljuc not in vidjeno:
            vidjeno.add(kljuc)
            rez.append(dict(p, ocjena=round(min(o, 1.0), 2)))
        if len(rez) == 5:
            break
    return rez

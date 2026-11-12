"""Popunjavanje Excel narudžbenica nakladnika (Školska knjiga, Profil Klett, Alfa...) količinama iz programa.

Obrazac (.xlsx) ostaje kakav je - mijenjaju se samo ćelije s količinom u redovima čiji se barkod nalazi
u narudžbenici. Ako obrazac nema stupac za količinu, doda se stupac "KOLIČINA" iza zadnjeg stupca.
Količine koje su već bile upisane u obrascu (a nisu u ovoj narudžbenici) brišu se.
Radi bez vanjskih biblioteka: xlsx je zip s XML datotekama koje se mijenjaju na mjestu.
"""
import html
import io
import re
import unicodedata
import xml.etree.ElementTree as ET
import zipfile
from decimal import Decimal, InvalidOperation

NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
BARKOD = ('BAR CODE', 'BARCODE', 'BARKOD', 'BAR KOD', 'EAN', 'ISBN')
KOLICINA = ('KOLICINA', 'KOL.', 'KOL', 'KOM', 'KOM.', 'NARUCENO', 'BROJ PRIMJERAKA')
CELIJA_RE = re.compile(r'<c\b([^>]*?)(?:/>|>(.*?)</c>)', re.S)
RED_RE = re.compile(r'<row\b([^>]*)>(.*?)</row>', re.S)


def bez_kvacica(s):
    return ''.join(c for c in unicodedata.normalize('NFD', str(s or '').upper()) if unicodedata.category(c) != 'Mn').strip()


def stupac_broj(slova):
    n = 0
    for z in slova:
        n = n * 26 + ord(z) - 64
    return n


def stupac_slova(n):
    s = ''
    while n:
        n, o = divmod(n - 1, 26)
        s = chr(65 + o) + s
    return s


def barkod_iz(v):
    """'9789530138643', '9.789530138643E12', '3859892042745.0' -> '9789530138643'."""
    v = str(v or '').strip()
    if re.fullmatch(r'[\d.,Ee+]+', v) and ('E' in v.upper() or '.' in v):
        try:
            v = str(int(Decimal(v.replace(',', '.'))))
        except (InvalidOperation, ValueError):
            pass
    d = re.sub(r'\D', '', v)
    return d if len(d) >= 8 else ''


def _zajednicki(z):
    try:
        root = ET.fromstring(z.read('xl/sharedStrings.xml'))
    except KeyError:
        return []
    return [''.join(t.text or '' for t in si.iter(NS + 't')) for si in root.iter(NS + 'si')]


def _listovi(z):
    """[(naziv lista, put do xml)] redom kao u radnoj knjizi."""
    wb = ET.fromstring(z.read('xl/workbook.xml'))
    rels = ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
    cilj = {r.get('Id'): r.get('Target') for r in rels}
    out = []
    for s in wb.iter(NS + 'sheet'):
        rid = s.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
        t = cilj.get(rid, '')
        t = t.lstrip('/') if t.startswith('/') else 'xl/' + t
        out.append((s.get('name'), t.replace('xl/xl/', 'xl/')))
    return out


def _vrijednost(atr, sadrzaj, ss):
    if not sadrzaj:
        return ''
    if 't="s"' in atr:
        m = re.search(r'<v>(\d+)</v>', sadrzaj)
        return ss[int(m.group(1))] if m else ''
    if 't="inlineStr"' in atr:
        return html.unescape(''.join(re.findall(r'<t[^>]*>(.*?)</t>', sadrzaj, re.S)))
    m = re.search(r'<v>(.*?)</v>', sadrzaj, re.S)
    return html.unescape(m.group(1)) if m else ''


def _celije(red_xml, ss):
    """{stupac: (atributi, sadrzaj, vrijednost, početak, kraj)} jednog reda."""
    out = {}
    for m in CELIJA_RE.finditer(red_xml):
        atr, sad = m.group(1), m.group(2) or ''
        r = re.search(r'\br="([A-Z]+)\d+"', atr)
        if r:
            out[r.group(1)] = (atr, sad, _vrijednost(atr, sad, ss), m.start(), m.end())
    return out


def _zaglavlje(celije):
    """Ako je red zaglavlje (ima stupac za barkod), vraća (stupac barkoda, stupac količine ili None)."""
    bk = kol = None
    for st, c in celije.items():
        t = bez_kvacica(c[2])
        if not t or len(t) > 40:
            continue
        if bk is None and any(t == x or t.startswith(x) for x in BARKOD):
            bk = st
        elif kol is None and any(t == x or (len(x) > 4 and t.startswith(x)) for x in KOLICINA):
            kol = st
    return (bk, kol) if bk else None


def _postavi(red_xml, red_br, stupac, nova_celija, celije):
    """Zamijeni ili umetni ćeliju u red (ćelije moraju ostati poredane po stupcu)."""
    if stupac in celije:
        _, _, _, a, b = celije[stupac]
        return red_xml[:a] + nova_celija + red_xml[b:]
    n = stupac_broj(stupac)
    for st, c in sorted(celije.items(), key=lambda x: stupac_broj(x[0])):
        if stupac_broj(st) > n:
            return red_xml[:c[3]] + nova_celija + red_xml[c[3]:]
    return red_xml + nova_celija


def _stil(atr):
    m = re.search(r'\bs="(\d+)"', atr)
    return ' s="%s"' % m.group(1) if m else ''


def popuni(sadrzaj, kolicine):
    """sadrzaj = bajtovi .xlsx obrasca, kolicine = {barkod: količina}.
    Vraća (bajtovi popunjenog obrasca, {barkod: upisano}, opis po listovima)."""
    if sadrzaj[:4] == b'\xd0\xcf\x11\xe0':
        raise ValueError('Ovo je stari Excel format (.xls). Otvorite obrazac u Excelu i spremite ga kao '
                         '"Excel radna knjiga (.xlsx)" (Datoteka -> Spremi kao), pa ga učitajte ponovno.')
    try:
        zin = zipfile.ZipFile(io.BytesIO(sadrzaj))
    except zipfile.BadZipFile:
        raise ValueError('Datoteka nije Excel (.xlsx) obrazac.')
    ss = _zajednicki(zin)
    kolicine = {barkod_iz(k): v for k, v in kolicine.items() if barkod_iz(k) and v}
    upisano, opis, promjene = {}, [], {}
    for naziv, put in _listovi(zin):
        try:
            xml = zin.read(put).decode('utf-8')
        except KeyError:
            continue
        stupci, dijelovi, zadnji, n_upisa, obrisano = None, [], 0, 0, 0
        for m in RED_RE.finditer(xml):
            atr, tijelo = m.group(1), m.group(2)
            rb = re.search(r'\br="(\d+)"', atr)
            if not rb:
                continue
            rb = rb.group(1)
            celije = _celije(tijelo, ss)
            novo = tijelo
            if stupci is None:
                z = _zaglavlje(celije)
                if z:
                    bk, kol = z
                    if kol is None:  # obrazac bez stupca za količinu (npr. Školska knjiga) - dodaj ga
                        kol = stupac_slova(max(stupac_broj(s) for s in celije) + 1)
                        ref = celije[bk][0]
                        novo = _postavi(tijelo, rb, kol, '<c r="%s%s"%s t="inlineStr"><is><t>KOLIČINA</t></is></c>'
                                        % (kol, rb, _stil(ref)), celije)
                    stupci = (bk, kol)
            else:
                bk, kol = stupci
                b = barkod_iz(celije.get(bk, ('', '', ''))[2])
                if b and b in kolicine:
                    k = kolicine[b]
                    stil = _stil(celije[kol][0]) if kol in celije else _stil(celije[bk][0])
                    novo = _postavi(tijelo, rb, kol, '<c r="%s%s"%s><v>%s</v></c>' % (kol, rb, stil, k), celije)
                    upisano[b] = upisano.get(b, 0) + k
                    n_upisa += 1
                elif kol in celije and '<f' not in celije[kol][1] and str(celije[kol][2]).strip():
                    # stara količina koja je ostala u obrascu (npr. od prošle narudžbe) - obriši je
                    novo = _postavi(tijelo, rb, kol, '<c r="%s%s"%s/>' % (kol, rb, _stil(celije[kol][0])), celije)
                    obrisano += 1
            if novo is not tijelo:
                atr = re.sub(r'\s+spans="[^"]*"', '', atr)  # spans više ne odgovara ako je dodan stupac
                dijelovi.append((m.start(), m.end(), '<row%s>%s</row>' % (atr, novo)))
        if dijelovi:
            for a, b, t in reversed(dijelovi):
                xml = xml[:a] + t + xml[b:]
            promjene[put] = xml.encode('utf-8')
        if stupci:
            opis.append({'list': naziv, 'upisano': n_upisa, 'obrisano': obrisano})
    if not opis:
        raise ValueError('U obrascu nisam našao stupac s barkodom (npr. "Bar code", "EAN" ili "Barkod").')
    # Excel neka sve formule (iznosi, zbrojevi) preračuna kod otvaranja
    wb = zin.read('xl/workbook.xml').decode('utf-8')
    if '<calcPr' in wb:
        wb = re.sub(r'<calcPr\b(?![^>]*fullCalcOnLoad)', '<calcPr fullCalcOnLoad="1"', wb, 1)
    else:
        wb = wb.replace('</workbook>', '<calcPr fullCalcOnLoad="1"/></workbook>')
    promjene['xl/workbook.xml'] = wb.encode('utf-8')
    izlaz = io.BytesIO()
    with zipfile.ZipFile(izlaz, 'w', zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            if info.filename == 'xl/calcChain.xml':
                continue  # Excel ga napravi ponovno
            d = promjene.get(info.filename) or zin.read(info.filename)
            if info.filename == '[Content_Types].xml':
                d = re.sub(rb'<Override[^>]*calcChain[^>]*/>', b'', d)
            elif info.filename == 'xl/_rels/workbook.xml.rels':
                d = re.sub(rb'<Relationship[^>]*calcChain[^>]*/>', b'', d)
            nova = zipfile.ZipInfo(info.filename, info.date_time)  # nova ZipInfo - original se ne smije mijenjati
            nova.compress_type = zipfile.ZIP_DEFLATED
            nova.external_attr = info.external_attr
            zout.writestr(nova, d)
    return izlaz.getvalue(), upisano, opis

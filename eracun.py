"""Čitanje eRačuna nakladnika (UBL 2.1 XML, Fiskalizacija 2.0) i otpremnica u Excelu - za "Stiglo".

Iz eRačuna se uzimaju stavke: naziv, barkod (GTIN/EAN), šifra prodavatelja (npr. šifra Školske knjige), količina, cijena.
"""
import io
import re
import xml.etree.ElementTree as ET
import zipfile

import popis


def _ime(el):
    return el.tag.rsplit('}', 1)[-1]


def _dijete(el, *put):
    for korak in put:
        if el is None:
            return None
        el = next((d for d in el if _ime(d) == korak), None)
    return el


def _tekst(el, *put):
    e = _dijete(el, *put)
    return (e.text or '').strip() if e is not None and e.text else ''


def _broj(t):
    try:
        return float(str(t).replace(',', '.'))
    except (TypeError, ValueError):
        return 0.0


def citaj_ubl(xml_bajtovi):
    try:
        korijen = ET.fromstring(xml_bajtovi)
    except ET.ParseError as e:
        raise ValueError('XML datoteka nije ispravna: %s' % e)
    # eRačun ponekad dolazi "umotan" (npr. kod posrednika) - potraži Invoice / CreditNote bilo gdje
    racun = korijen if _ime(korijen) in ('Invoice', 'CreditNote') else \
        next((e for e in korijen.iter() if _ime(e) in ('Invoice', 'CreditNote')), None)
    if racun is None:
        raise ValueError('U XML datoteci nije eRačun (nema Invoice).')
    odobrenje = _ime(racun) == 'CreditNote'
    dobavljac = _tekst(racun, 'AccountingSupplierParty', 'Party', 'PartyLegalEntity', 'RegistrationName') or \
        _tekst(racun, 'AccountingSupplierParty', 'Party', 'PartyName', 'Name')
    stavke = []
    for red in racun:
        if _ime(red) not in ('InvoiceLine', 'CreditNoteLine'):
            continue
        kol_el = _dijete(red, 'InvoicedQuantity')
        if kol_el is None:
            kol_el = _dijete(red, 'CreditedQuantity')
        item = _dijete(red, 'Item')
        naziv = _tekst(item, 'Name') or _tekst(item, 'Description')
        barkod = ''
        for e in (item if item is not None else []):
            if _ime(e) == 'StandardItemIdentification':
                barkod = re.sub(r'\D', '', _tekst(e, 'ID'))
        sifra = _tekst(item, 'SellersItemIdentification', 'ID')
        if not barkod:  # neki stavljaju EAN u šifru prodavatelja ili u naziv
            for kand in (sifra, _tekst(item, 'BuyersItemIdentification', 'ID'), naziv):
                m = re.search(r'(?<!\d)(97[89]\d{10}|38[0-9]\d{10})(?!\d)', kand or '')
                if m:
                    barkod = m.group(1)
                    break
        stavke.append({'naziv': naziv, 'barkod': barkod if len(barkod) in (8, 12, 13, 14) else '',
                       'sifra': sifra, 'kolicina': _broj(kol_el.text if kol_el is not None else 0),
                       'cijena': _broj(_tekst(red, 'Price', 'PriceAmount'))})
    return {'broj': _tekst(racun, 'ID'), 'datum': _tekst(racun, 'IssueDate'), 'dobavljac': dobavljac,
            'odobrenje': odobrenje, 'stavke': stavke}


def citaj_otpremnicu(ime, sadrzaj):
    """eRačun (.xml, ili .zip s XML-om) ili otpremnica/račun u Excelu/CSV-u sa stupcima barkod i količina."""
    if sadrzaj[:2] == b'PK' and not ime.lower().endswith(('.xlsx', '.docx')):
        z = zipfile.ZipFile(io.BytesIO(sadrzaj))
        xmlovi = [n for n in z.namelist() if n.lower().endswith('.xml')]
        if xmlovi:
            return citaj_ubl(z.read(xmlovi[0]))
    if ime.lower().endswith('.xml') or sadrzaj.lstrip()[:5] in (b'<?xml', b'<Invo', b'<Cred') or \
            sadrzaj.lstrip(b'\xef\xbb\xbf \r\n\t')[:1] == b'<' and b'Invoice' in sadrzaj[:5000]:
        return citaj_ubl(sadrzaj)
    redovi = popis.redovi_iz_datoteke(ime, sadrzaj)
    stavke = [s for s in popis.narudzba_iz_redova(redovi) if s['kolicina'] > 0]
    if not stavke:
        raise ValueError('U datoteci nisam našao tablicu s nazivom/barkodom i količinom.')
    return {'broj': '', 'datum': '', 'dobavljac': '', 'odobrenje': False,
            'stavke': [{'naziv': s['naslov'], 'barkod': s['barkod'], 'sifra': s['sifra'], 'kolicina': s['kolicina'],
                        'cijena': s['cijena']} for s in stavke]}

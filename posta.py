"""Slanje e-maila s prilogom izravno iz programa (SMTP, npr. Gmail) i izrada jednostavne Excel (.xlsx) narudžbenice.

Postavke e-maila (poslužitelj, korisnik, lozinka) čuvaju se u postavke.json u mapi programa, NE u bazi,
da lozinka ne završi u kopijama baze ili Setupu.
"""
import datetime
import io
import mimetypes
import smtplib
import ssl
import zipfile
from email.message import EmailMessage
from email.utils import formataddr, make_msgid
from xml.sax.saxutils import escape

POSLUZITELJI = {  # poznati poslužitelji po domeni adrese
    'gmail.com': ('smtp.gmail.com', 587), 'googlemail.com': ('smtp.gmail.com', 587),
    'outlook.com': ('smtp-mail.outlook.com', 587), 'hotmail.com': ('smtp-mail.outlook.com', 587),
    'live.com': ('smtp-mail.outlook.com', 587), 'yahoo.com': ('smtp.mail.yahoo.com', 587),
    't-com.hr': ('mail.t-com.hr', 587), 'net.hr': ('smtp.net.hr', 587), 'inet.hr': ('mail.inet.hr', 587),
    'gmx.com': ('mail.gmx.com', 587), 'gmx.net': ('mail.gmx.net', 587), 'icloud.com': ('smtp.mail.me.com', 587),
}


def zadani_posluzitelj(adresa):
    return POSLUZITELJI.get((adresa or '').rsplit('@', 1)[-1].lower().strip(), ('', 587))


def posalji(cfg, prima, naslov, tijelo, prilozi=(), kopija_sebi=True):
    """cfg = {'adresa', 'lozinka', 'posluzitelj', 'port', 'ime'}; prilozi = [(ime datoteke, bajtovi)]."""
    adresa = (cfg.get('adresa') or '').strip()
    if not adresa or not cfg.get('lozinka'):
        raise ValueError('E-mail nije podešen: u Postavkama upišite svoju e-mail adresu i lozinku za aplikacije.')
    posluzitelj = (cfg.get('posluzitelj') or '').strip() or zadani_posluzitelj(adresa)[0]
    port = int(cfg.get('port') or zadani_posluzitelj(adresa)[1] or 587)
    if not posluzitelj:
        raise ValueError('Za adresu %s ne znam poslužitelj - upišite ga u Postavkama (npr. smtp.gmail.com).' % adresa)
    prima = [p.strip() for p in str(prima or '').replace(';', ',').split(',') if '@' in p]
    if not prima:
        raise ValueError('Nije upisana e-mail adresa primatelja.')
    m = EmailMessage()
    m['From'] = formataddr((cfg.get('ime') or '', adresa))
    m['To'] = ', '.join(prima)
    if kopija_sebi:
        m['Bcc'] = adresa  # kopija u vlastiti sandučić, da se vidi što je poslano
    m['Subject'] = naslov
    m['Message-ID'] = make_msgid()
    m.set_content(tijelo)
    for ime, bajtovi in prilozi:
        tip = (mimetypes.guess_type(ime)[0] or 'application/octet-stream').split('/', 1)
        m.add_attachment(bajtovi, maintype=tip[0], subtype=tip[1], filename=ime)
    kontekst = ssl.create_default_context()
    try:
        if port == 465:
            with smtplib.SMTP_SSL(posluzitelj, port, context=kontekst, timeout=30) as s:
                s.login(adresa, cfg['lozinka'].replace(' ', '') if 'gmail' in posluzitelj else cfg['lozinka'])
                s.send_message(m)
        else:
            with smtplib.SMTP(posluzitelj, port, timeout=30) as s:
                s.ehlo()
                if s.has_extn('starttls'):
                    s.starttls(context=kontekst)
                    s.ehlo()
                if s.has_extn('auth'):
                    s.login(adresa, cfg['lozinka'].replace(' ', '') if 'gmail' in posluzitelj else cfg['lozinka'])
                s.send_message(m)
    except smtplib.SMTPAuthenticationError:
        raise ValueError('Prijava na e-mail nije uspjela. Za Gmail treba "lozinka za aplikacije" (16 slova), '
                         'ne obična lozinka - vidi upute u Postavkama.')
    except (OSError, smtplib.SMTPException) as e:
        raise ValueError('Slanje e-maila nije uspjelo: %s' % e)
    return prima


# ------------------------------------------------------------- jednostavan .xlsx

def _stupac(n):
    s = ''
    n += 1
    while n:
        n, o = divmod(n - 1, 26)
        s = chr(65 + o) + s
    return s


def xlsx(redovi, sirine=(), podebljani=()):
    """redovi = lista listi (str / int / float). Brojevi ostaju brojevi, a tekst (i barkod) tekst."""
    xml_redovi = []
    for r, red in enumerate(redovi):
        celije = []
        for s, v in enumerate(red):
            ref = '%s%d' % (_stupac(s), r + 1)
            stil = ' s="1"' if r in podebljani else ''
            if v is None or v == '':
                continue
            if isinstance(v, (int, float)) and not isinstance(v, bool):
                celije.append('<c r="%s"%s><v>%s</v></c>' % (ref, stil, v))
            else:
                celije.append('<c r="%s"%s t="inlineStr"><is><t xml:space="preserve">%s</t></is></c>'
                              % (ref, stil, escape(str(v))))
        xml_redovi.append('<row r="%d">%s</row>' % (r + 1, ''.join(celije)))
    stupci = ''.join('<col min="%d" max="%d" width="%s" customWidth="1"/>' % (i + 1, i + 1, w) for i, w in enumerate(sirine))
    list_ = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
             '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
             + ('<cols>%s</cols>' % stupci if stupci else '') + '<sheetData>%s</sheetData></worksheet>' % ''.join(xml_redovi))
    datoteke = {
        '[Content_Types].xml': '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
        '</Types>',
        '_rels/.rels': '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
        '</Relationships>',
        'xl/workbook.xml': '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<sheets><sheet name="Narudzbenica" sheetId="1" r:id="rId1"/></sheets></workbook>',
        'xl/_rels/workbook.xml.rels': '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
        '</Relationships>',
        'xl/styles.xml': '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts>'
        '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>'
        '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        '<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
        '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/></cellXfs>'
        '</styleSheet>',
        'xl/worksheets/sheet1.xml': list_,
    }
    izlaz = io.BytesIO()
    with zipfile.ZipFile(izlaz, 'w', zipfile.ZIP_DEFLATED) as z:
        for ime, sad in datoteke.items():
            z.writestr(zipfile.ZipInfo(ime, datetime.datetime.now().timetuple()[:6]), sad.encode('utf-8'))
    return izlaz.getvalue()

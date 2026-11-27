"""Čitanje starih Office formata bez vanjskih biblioteka: Excel 97-2003 (.xls) i Word 97-2003 (.doc).

Obje datoteke su "OLE2 compound file" (mali datotečni sustav u datoteci). Iz .xls se čitaju vrijednosti ćelija
svih listova (redovi kao kod .xlsx), a iz .doc tekst s tablicama (red tablice = red s ćelijama).
"""
import struct

OLE_POTPIS = b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1'
KRAJ = 0xFFFFFFFE


class Ole:
    def __init__(self, podaci):
        if podaci[:8] != OLE_POTPIS:
            raise ValueError('Datoteka nije u starom Office formatu.')
        self.p = podaci
        self.sektor = 1 << struct.unpack_from('<H', podaci, 0x1E)[0]
        self.mali_sektor = 1 << struct.unpack_from('<H', podaci, 0x20)[0]
        n_fat, prvi_dir = struct.unpack_from('<II', podaci, 0x2C)
        self.granica = struct.unpack_from('<I', podaci, 0x38)[0]
        prvi_minifat, n_minifat, prvi_difat, n_difat = struct.unpack_from('<IIII', podaci, 0x3C)
        difat = list(struct.unpack_from('<109I', podaci, 0x4C))
        s = prvi_difat
        for _ in range(n_difat):
            if s >= KRAJ - 1:
                break
            blok = self._sektor(s)
            vr = struct.unpack('<%dI' % (self.sektor // 4), blok)
            difat += vr[:-1]
            s = vr[-1]
        self.fat = []
        for s in difat[:n_fat]:
            if s < KRAJ - 1:
                self.fat += struct.unpack('<%dI' % (self.sektor // 4), self._sektor(s))
        dirp = self._lanac(prvi_dir)
        self.unosi = {}
        korijen = None
        for i in range(len(dirp) // 128):
            e = dirp[i * 128:(i + 1) * 128]
            duz = struct.unpack_from('<H', e, 64)[0]
            ime = e[:max(0, duz - 2)].decode('utf-16-le', 'replace')
            tip = e[66]
            start, vel = struct.unpack_from('<II', e, 116)
            if tip == 5:
                korijen = (start, vel)
            elif tip == 2:
                self.unosi.setdefault(ime, (start, vel))
        self.mini = self._lanac(korijen[0])[:korijen[1]] if korijen and korijen[0] < KRAJ - 1 else b''
        self.minifat = []
        if n_minifat and prvi_minifat < KRAJ - 1:
            mf = self._lanac(prvi_minifat)
            self.minifat = list(struct.unpack('<%dI' % (len(mf) // 4), mf))

    def _sektor(self, s):
        a = (s + 1) * self.sektor
        return self.p[a:a + self.sektor]

    def _lanac(self, s):
        out, vidjeno = [], set()
        while s < KRAJ - 1 and s < len(self.fat) and s not in vidjeno:
            vidjeno.add(s)
            out.append(self._sektor(s))
            s = self.fat[s]
        return b''.join(out)

    def tok(self, ime):
        if ime not in self.unosi:
            return None
        start, vel = self.unosi[ime]
        if vel < self.granica:
            out, s, vidjeno = [], start, set()
            while s < KRAJ - 1 and s < len(self.minifat) and s not in vidjeno:
                vidjeno.add(s)
                a = s * self.mali_sektor
                out.append(self.mini[a:a + self.mali_sektor])
                s = self.minifat[s]
            return b''.join(out)[:vel]
        return self._lanac(start)[:vel]


# ------------------------------------------------------------- Excel .xls (BIFF8)

def _broj(v):
    if isinstance(v, float) and v.is_integer() and abs(v) < 1e15:
        return str(int(v))
    return str(v)


def _rk(rk):
    if rk & 2:
        v = rk >> 2
        if v & 0x20000000:
            v -= 0x40000000
        v = float(v)
    else:
        v = struct.unpack('<d', struct.pack('<Q', (rk & 0xFFFFFFFC) << 32))[0]
    return v / 100 if rk & 1 else v


def _zapisi(tok):
    i = 0
    while i + 4 <= len(tok):
        oznaka, duz = struct.unpack_from('<HH', tok, i)
        yield oznaka, tok[i + 4:i + 4 + duz]
        i += 4 + duz


def _sst(dijelovi):
    """Tablica zajedničkih tekstova; tekst se može prelomiti preko CONTINUE zapisa (tada dolazi nova oznaka kodiranja)."""
    out = []
    if not dijelovi:
        return out
    di, pos = 0, 8
    ukupno = struct.unpack_from('<I', dijelovi[0], 4)[0]

    def citaj(n):
        nonlocal di, pos
        b = b''
        while n:
            if pos >= len(dijelovi[di]):
                di += 1
                pos = 0
            uzmi = min(n, len(dijelovi[di]) - pos)
            b += dijelovi[di][pos:pos + uzmi]
            pos += uzmi
            n -= uzmi
        return b

    try:
        for _ in range(ukupno):
            if pos >= len(dijelovi[di]):
                di, pos = di + 1, 0
            znakova = struct.unpack('<H', citaj(2))[0]
            zast = citaj(1)[0]
            rich = struct.unpack('<H', citaj(2))[0] if zast & 8 else 0
            ext = struct.unpack('<I', citaj(4))[0] if zast & 4 else 0
            siroko, tekst, ostalo = zast & 1, [], znakova
            while ostalo:
                if pos >= len(dijelovi[di]):  # nastavak u CONTINUE: prvi bajt je nova oznaka (1 ili 2 bajta po znaku)
                    di, pos = di + 1, 0
                    siroko = dijelovi[di][0] & 1
                    pos = 1
                stane = (len(dijelovi[di]) - pos) // (2 if siroko else 1)
                n = min(ostalo, stane)
                b = dijelovi[di][pos:pos + n * (2 if siroko else 1)]
                pos += len(b)
                tekst.append(b.decode('utf-16-le' if siroko else 'latin-1', 'replace'))
                ostalo -= n
            citaj(rich * 4 + ext)
            out.append(''.join(tekst))
    except (IndexError, struct.error):
        pass
    return out


def _xl_tekst(d, pos, duzina_bajtova=2):
    n = struct.unpack_from('<H' if duzina_bajtova == 2 else '<B', d, pos)[0]
    pos += duzina_bajtova
    zast = d[pos]
    pos += 1
    if zast & 1:
        return d[pos:pos + 2 * n].decode('utf-16-le', 'replace')
    return d[pos:pos + n].decode('latin-1', 'replace')


def citaj_xls(podaci):
    ole = Ole(podaci)
    tok = ole.tok('Workbook') or ole.tok('Book')
    if not tok:
        raise ValueError('U .xls datoteci nema radne knjige.')
    zapisi = list(_zapisi(tok))
    sst_dijelovi, sst = [], []
    for i, (oz, d) in enumerate(zapisi):
        if oz == 0xFC:
            sst_dijelovi = [d]
            j = i + 1
            while j < len(zapisi) and zapisi[j][0] == 0x3C:
                sst_dijelovi.append(zapisi[j][1])
                j += 1
            break
    sst = _sst(sst_dijelovi)
    listovi, celije, u_listu, ceka_tekst = [], None, False, None
    for oz, d in zapisi:
        if oz == 0x809:  # BOF
            u_listu = len(d) >= 4 and struct.unpack_from('<H', d, 2)[0] == 0x10
            if u_listu:
                celije = {}
                listovi.append(celije)
            continue
        if not u_listu or celije is None:
            continue
        try:
            if oz == 0xFD:  # LABELSST
                r, s, _, idx = struct.unpack_from('<HHHI', d)
                celije[(r, s)] = sst[idx] if idx < len(sst) else ''
            elif oz == 0x203:  # NUMBER
                r, s, _, v = struct.unpack_from('<HHHd', d)
                celije[(r, s)] = _broj(v)
            elif oz == 0x27E:  # RK
                r, s, _, rk = struct.unpack_from('<HHHI', d)
                celije[(r, s)] = _broj(_rk(rk))
            elif oz == 0xBD:  # MULRK
                r, prvi = struct.unpack_from('<HH', d)
                n = (len(d) - 6) // 6
                for k in range(n):
                    rk = struct.unpack_from('<I', d, 4 + k * 6 + 2)[0]
                    celije[(r, prvi + k)] = _broj(_rk(rk))
            elif oz == 0x204:  # LABEL
                r, s = struct.unpack_from('<HH', d)
                celije[(r, s)] = _xl_tekst(d, 6)
            elif oz == 0x06:  # FORMULA - spremljeni rezultat
                r, s = struct.unpack_from('<HH', d)
                if d[12:14] == b'\xff\xff':
                    if d[6] == 0:
                        ceka_tekst = (r, s)
                    elif d[6] == 1:
                        celije[(r, s)] = 'TRUE' if d[8] else 'FALSE'
                else:
                    celije[(r, s)] = _broj(struct.unpack_from('<d', d, 6)[0])
            elif oz == 0x207 and ceka_tekst:  # STRING (tekstualni rezultat formule)
                celije[ceka_tekst] = _xl_tekst(d, 0)
                ceka_tekst = None
            elif oz == 0x205:  # BOOLERR
                r, s, _, v, je_greska = struct.unpack_from('<HHHBB', d)
                if not je_greska:
                    celije[(r, s)] = 'TRUE' if v else 'FALSE'
        except struct.error:
            continue
    redovi = []
    for c in listovi:
        if not c:
            continue
        for r in sorted({k[0] for k in c}):
            stupci = [k[1] for k in c if k[0] == r]
            redovi.append([c.get((r, s), '') for s in range(max(stupci) + 1)])
    return redovi


# ------------------------------------------------------------- Word .doc

def citaj_doc(podaci):
    ole = Ole(podaci)
    wd = ole.tok('WordDocument')
    if not wd or len(wd) < 0x1AA:
        raise ValueError('U .doc datoteci nema teksta.')
    zast = struct.unpack_from('<H', wd, 0x0A)[0]
    if zast & 0x0100:
        raise ValueError('Word dokument je zaštićen lozinkom.')
    tablica = ole.tok('1Table' if zast & 0x0200 else '0Table') or b''
    fc_clx, lcb_clx = struct.unpack_from('<II', wd, 0x1A2)
    clx = tablica[fc_clx:fc_clx + lcb_clx]
    i, tekst = 0, []
    while i < len(clx):
        if clx[i] == 1:  # Prc - preskoči
            i += 3 + struct.unpack_from('<H', clx, i + 1)[0]
        elif clx[i] == 2:  # Pcdt
            duz = struct.unpack_from('<I', clx, i + 1)[0]
            plc = clx[i + 5:i + 5 + duz]
            n = (duz - 4) // 12
            cp = struct.unpack_from('<%dI' % (n + 1), plc)
            for k in range(n):
                fc = struct.unpack_from('<I', plc, (n + 1) * 4 + k * 8 + 2)[0]
                znakova = cp[k + 1] - cp[k]
                if fc & 0x40000000:
                    a = (fc & 0x3FFFFFFF) // 2
                    tekst.append(wd[a:a + znakova].decode('cp1252', 'replace'))
                else:
                    tekst.append(wd[fc:fc + 2 * znakova].decode('utf-16-le', 'replace'))
            break
        else:
            break
    if not tekst:  # vrlo stari / "brzo spremljeni" dokumenti bez tablice dijelova: tekst je u jednom komadu
        fc_min, fc_mac = struct.unpack_from('<II', wd, 0x18)
        tekst = [wd[fc_min:fc_mac].decode('cp1252', 'replace')]
    return _doc_redovi(''.join(tekst))


def _doc_redovi(t):
    # polja: \x13 kod \x14 rezultat \x15 -> ostavi samo rezultat
    out, dubina, u_kodu = [], 0, []
    for z in t:
        if z == '\x13':
            dubina += 1
            u_kodu.append(True)
        elif z == '\x14':
            if u_kodu:
                u_kodu[-1] = False
        elif z == '\x15':
            if u_kodu:
                u_kodu.pop()
            dubina = max(0, dubina - 1)
        elif not any(u_kodu):
            out.append(z)
    t = ''.join(out)
    redovi, celije, trenutna, kraj_celije, stupaca = [], [], '', False, 0
    for z in t:
        if z == '\x07':
            if kraj_celije and not trenutna.strip() and not (stupaca and len(celije) < stupaca):
                # dvije oznake zaredom = kraj reda tablice (osim ako red još nema sve ćelije - onda je prazna ćelija)
                if celije:
                    redovi.append(celije)
                    stupaca = len(celije)
                celije, trenutna, kraj_celije = [], '', False
            else:
                celije.append(trenutna.strip())
                trenutna, kraj_celije = '', True
        elif z in '\r\x0c':
            if celije:
                trenutna += ' '  # novi odlomak unutar ćelije tablice
            else:
                if trenutna.strip():
                    redovi.append([trenutna.strip()])
                    stupaca = 0  # tekst izvan tablice - sljedeća tablica može imati drugi broj stupaca
                trenutna = ''
        else:
            trenutna += ' ' if z in '\x0b\t\xa0' else (z if z >= ' ' else '')
            if z.strip():
                kraj_celije = False
    if celije:
        redovi.append(celije)
    if trenutna.strip():
        redovi.append([trenutna.strip()])
    return [[' '.join(c.split()) for c in r] for r in redovi]

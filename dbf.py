"""Jednostavan čitač FoxPro / dBase DBF tablica (samo čitanje, bez vanjskih biblioteka)."""
import datetime
import os
import struct


def _fpt_reader(path):
    for ext in ('.fpt', '.FPT', '.Fpt'):
        p = os.path.splitext(path)[0] + ext
        if os.path.exists(p):
            f = open(p, 'rb')
            block = struct.unpack('>H', f.read(8)[6:8])[0] or 64
            return f, block
    return None, 0


def citaj(path, encoding='cp1250'):
    """Vraća listu rječnika {POLJE: vrijednost}; obrisani zapisi se preskaču."""
    with open(path, 'rb') as f:
        data = f.read()
    nrec, hlen, rlen = struct.unpack('<IHH', data[4:12])
    fields = []
    pos, off = 32, 1
    while pos < hlen and data[pos] != 0x0D:
        name = data[pos:pos + 11].split(b'\0')[0].decode('ascii', 'replace')
        ftype = chr(data[pos + 11])
        flen, fdec = data[pos + 16], data[pos + 17]
        fields.append((name, ftype, off, flen, fdec))
        off += flen
        pos += 32
    memo, block = (None, 0)
    if any(t in 'MG' for _, t, _, _, _ in fields):
        memo, block = _fpt_reader(path)
    out = []
    try:
        for i in range(nrec):
            rec = data[hlen + i * rlen: hlen + (i + 1) * rlen]
            if len(rec) < rlen or rec[0:1] == b'*':
                continue
            row = {}
            for name, t, o, ln, dec in fields:
                if name.startswith('_'):
                    continue
                raw = rec[o:o + ln]
                row[name] = _vrijednost(t, raw, dec, encoding, memo, block)
            out.append(row)
    finally:
        if memo:
            memo.close()
    return out


def _vrijednost(t, raw, dec, enc, memo, block):
    if t in 'CV':
        return raw.rstrip(b'\0 ').decode(enc, 'replace').strip()
    if t in 'NF':
        s = raw.strip(b' \0')
        if not s or s == b'.':
            return 0
        try:
            return float(s) if (dec or b'.' in s) else int(s)
        except ValueError:
            return 0
    if t == 'I':
        return struct.unpack('<i', raw)[0]
    if t == 'B':
        return struct.unpack('<d', raw)[0]
    if t == 'Y':
        return struct.unpack('<q', raw)[0] / 10000.0
    if t == 'L':
        return raw in (b'T', b't', b'Y', b'y')
    if t == 'D':
        s = raw.decode('ascii', 'replace').strip()
        try:
            return datetime.date(int(s[:4]), int(s[4:6]), int(s[6:8]))
        except ValueError:
            return None
    if t == 'M' and memo:
        n = struct.unpack('<i', raw)[0] if len(raw) == 4 else int(raw.strip() or 0)
        if n <= 0:
            return ''
        memo.seek(n * block)
        _, ln = struct.unpack('>II', memo.read(8))
        return memo.read(ln).rstrip(b'\0').decode(enc, 'replace').strip()
    return None

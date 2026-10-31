#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Debitor-Kreditor proqramı
=========================
- Qaimələr   -> qaimeler/   papkası (şablon: sablonlar/sablon_qaime.xlsx)
- Ödənişlər  -> odenisler/  papkası (şablon: sablonlar/sablon_odenis.xlsx)
- "Papkaları yenilə" düyməsi yalnız YENİ və ya DƏYİŞMİŞ faylları oxuyur.
- Məlumat bazası: SQLite (tək fayl, quraşdırma tələb etmir). Məbləğlər qəpiklə
  (tam ədəd) saxlanır - toplamalarda yuvarlaqlaşdırma xətası olmur.

BİRLƏŞDİRİLMİŞ TƏK FAYL: debitor-kreditor + veb ERP (yaşlandırma, risk/limit, ehtiyat) +
Mənfəət-zərər modulu (Excel iş kitabından) + Chrome uzantısı "Vergi alətləri" v1.6 (içində).
Ayrıca erp_web.py lazım deyil.

İstifadə:
    python debitor_kreditor.py --maza-import maza_data --maza-export hesabat.xlsx --il 2026   # M.A.Z.A
    python debitor_kreditor.py --backup             # bazanın ehtiyat nüsxəsi
    python debitor_kreditor.py --uzanti             # Chrome uzantısını uzanti qovluğuna çıxar
    python debitor_kreditor.py                      # pəncərəli proqram
    python debitor_kreditor.py --sablon             # Excel şablonlarını yarat
    python debitor_kreditor.py --yenile             # papkaları oxu
    python debitor_kreditor.py --export hesabat.xlsx [--tarix 31.12.2026]
"""
import io
import os
import re
import sys
import base64
import secrets
import tempfile
import webbrowser
from urllib.parse import urlparse, parse_qs
import json
import threading
import urllib.request
import urllib.error
import time
import sqlite3
import argparse
import datetime as dt
from decimal import Decimal, ROUND_HALF_UP, InvalidOperation
from pathlib import Path

for _n in ("stdout", "stderr"):      # pəncərəli (konsolsuz) exe-də None ola bilər
    if getattr(sys, _n) is None:
        setattr(sys, _n, open(os.devnull, "w"))
    try:
        getattr(sys, _n).reconfigure(errors="replace")      # Windows konsolunda Unicode xətası olmasın
    except Exception:
        pass

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.utils.datetime import from_excel
from openpyxl.worksheet.datavalidation import DataValidation

def _base_dir():
    """Məlumatlar (baza, config, papkalar) daimi yerdə saxlanır.
    .exe (PyInstaller) müvəqqəti qovluğa açılır və bağlananda silinir - ona görə __file__ yox,
    .exe-nin öz qovluğu götürülür. Yazmaq olmursa: %APPDATA%\\DebitorKreditor."""
    if getattr(sys, "frozen", False):
        d = Path(sys.executable).resolve().parent
    else:
        d = Path(__file__).resolve().parent
    try:
        t = d / ".yaz_test"
        t.write_text("1")
        t.unlink()
        return d
    except Exception:
        import os
        alt = Path(os.environ.get("APPDATA") or Path.home()) / "DebitorKreditor"
        alt.mkdir(parents=True, exist_ok=True)
        return alt


BASE = _base_dir()
CONFIG = BASE / "config.json"
VAT_DEFAULT = 18  # ƏDV dərəcəsi (%), şablonda "ƏDV %" sütunu boşdursa
DEFAULTS = {
    "sessionId": "",   # lazım olsa: portalda giriş edilmiş sessiya (adətən boş qalır)
    "qaime": str(BASE / "qaimeler"),
    "odenis": str(BASE / "odenisler"),
    "db": str(BASE / "debitor_kreditor.db"),
}

KIND_LABEL = {
    "S": "Satış qaiməsi",
    "SQ": "Satış qaytarılması",
    "A": "Alış qaiməsi",
    "AQ": "Alış qaytarılması",
    "D": "Daxil olan ödəniş",
    "DQ": "Müştəriyə pul qaytarılması",
    "CQ": "Təchizatçıdan pul qaytarılması",
    "C": "Çıxan ödəniş",
}

# ----------------------------------------------------------------------------
# Konfiqurasiya
# ----------------------------------------------------------------------------

def load_cfg():
    cfg = dict(DEFAULTS)
    if CONFIG.exists():
        try:
            cfg.update(json.loads(CONFIG.read_text(encoding="utf-8")))
        except Exception:
            pass
    return cfg


def save_cfg(cfg):
    CONFIG.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8")


# ----------------------------------------------------------------------------
# Verilənlər bazası (SQLite)
# ----------------------------------------------------------------------------

SCHEMA = """
CREATE TABLE IF NOT EXISTS files(path TEXT PRIMARY KEY, mtime REAL, size INTEGER);
CREATE TABLE IF NOT EXISTS parties(voen TEXT PRIMARY KEY, name TEXT, reg_name TEXT, vat TEXT, risky TEXT, active TEXT, checked TEXT);
CREATE TABLE IF NOT EXISTS txn(
    id INTEGER PRIMARY KEY,
    src TEXT NOT NULL,
    row_no INTEGER,
    kind TEXT NOT NULL,          -- S satış, SQ satış qaytarılma, A alış, AQ alış qaytarılma, D daxil olan, C çıxan
    date TEXT NOT NULL,          -- ISO: YYYY-MM-DD
    doc_no TEXT,
    voen TEXT NOT NULL,
    amount INTEGER NOT NULL,     -- qəpik (qaytarılmalar da müsbət saxlanır)
    note TEXT,
    vat INTEGER DEFAULT 18,      -- ƏDV % (yalnız 'part' boşdursa istifadə olunur)
    part TEXT DEFAULT ''         -- E = məbləğ əsasdır, V = məbləğ ƏDV-dir, boş = ƏDV daxil
);
CREATE INDEX IF NOT EXISTS ix_txn_voen_date ON txn(voen, date);
CREATE INDEX IF NOT EXISTS ix_txn_src ON txn(src);
"""


def connect(path):
    con = sqlite3.connect(str(path))
    con.execute("PRAGMA journal_mode=WAL")
    con.execute("PRAGMA synchronous=NORMAL")
    con.execute("PRAGMA temp_store=MEMORY")
    con.execute("PRAGMA cache_size=-65536")
    con.executescript(SCHEMA)
    cols = [r[1] for r in con.execute("PRAGMA table_info(txn)")]
    pcols = [r[1] for r in con.execute("PRAGMA table_info(parties)")]
    for c in ("reg_name", "vat", "risky", "active", "checked"):
        if c not in pcols:
            con.execute(f"ALTER TABLE parties ADD COLUMN {c} TEXT")
    if "part" not in cols:
        con.execute("ALTER TABLE txn ADD COLUMN part TEXT DEFAULT ''")
    if "vat" not in cols:
        con.execute("ALTER TABLE txn ADD COLUMN vat INTEGER DEFAULT 18")
    # köhnə bazada mənfi qaimələr = qaytarılma -> ayrı növə keçir
    with con:
        con.execute("UPDATE txn SET kind='SQ', amount=-amount WHERE kind='S' AND amount<0")
        con.execute("UPDATE txn SET kind='AQ', amount=-amount WHERE kind='A' AND amount<0")
    return con


# ----------------------------------------------------------------------------
# Dəyərlərin təhlili
# ----------------------------------------------------------------------------

_TR = str.maketrans("əıöüçşğ", "eioucsg")


def norm(s):
    """Azərbaycan hərflərini sadələşdirib kiçik hərfə çevirir (müqayisə üçün)."""
    return str(s).strip().lower().replace("i\u0307", "i").translate(_TR)


def parse_date(v):
    if v is None:
        return None
    if isinstance(v, dt.datetime):
        return v.date().isoformat()
    if isinstance(v, dt.date):
        return v.isoformat()
    if isinstance(v, (int, float)):
        if 30000 < v < 80000:
            return from_excel(v).date().isoformat()
        return None
    s = str(v).strip()
    for fmt in ("%d.%m.%Y", "%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%d.%m.%y"):
        try:
            return dt.datetime.strptime(s, fmt).date().isoformat()
        except ValueError:
            pass
    return None


def parse_amount(v):
    """Məbləği qəpiklə (int) qaytarır. '1 234,56' və '1,234.56' formatlarını anlayır."""
    if v is None or str(v).strip() == "":
        raise ValueError("məbləğ boşdur")
    if isinstance(v, (int, float)):
        d = Decimal(str(v))
    else:
        s = str(v).strip().replace(" ", "").replace("\u00a0", "")
        if "," in s and "." in s:
            if s.rfind(",") > s.rfind("."):
                s = s.replace(".", "").replace(",", ".")
            else:
                s = s.replace(",", "")
        elif "," in s:
            s = s.replace(",", ".")
        try:
            d = Decimal(s)
        except InvalidOperation:
            raise ValueError("məbləğ rəqəm deyil")
    return int((d * 100).to_integral_value(rounding=ROUND_HALF_UP))


def parse_vat(v):
    """ƏDV dərəcəsi (%). Boşdursa 18. '18', '18%', '0', 'ƏDV-siz' anlaşılır."""
    if v is None or str(v).strip() == "":
        return VAT_DEFAULT
    s = str(v).strip().replace("%", "").replace(",", ".")
    try:
        d = Decimal(s)
        if d < 1 and d > 0:   # Excel-də 18% = 0.18
            d *= 100
        return int(d)
    except InvalidOperation:
        return 0 if norm(s).startswith(("yox", "siz", "ed")) else VAT_DEFAULT


def vat_part(amount, rate):
    """ƏDV daxil məbləğdən ƏDV hissəsi (qəpiklə)."""
    if not rate:
        return 0
    return int((Decimal(amount) * rate / (100 + rate)).to_integral_value(rounding=ROUND_HALF_UP))


def parse_part(v):
    """'Əsas' -> 'E', 'ƏDV' -> 'V', digər -> ''."""
    if v is None or isinstance(v, (int, float)):
        return ""
    n = norm(v)
    if n.startswith("es"):
        return "E"
    if n.startswith("edv") or n.startswith("ed"):
        return "V"
    return ""


def split_amount(amount, part, rate):
    """(əsas, ƏDV) qəpiklə. part: E/V -> məbləğ tam o hissədir; boş -> ƏDV daxil sayılıb ayrılır."""
    if part == "E":
        return amount, 0
    if part == "V":
        return 0, amount
    vt = vat_part(amount, rate if rate is not None else 0)
    return amount - vt, vt


def parse_voen(v):
    if v is None:
        return None
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    s = str(v).strip()
    return s or None


def kind_of(ftype, s):
    n = norm(s or "")
    if ftype == "Q":
        ret = "qayt" in n or "iade" in n
        if n.startswith("sat"):
            return "SQ" if ret else "S"
        if n.startswith("al"):
            return "AQ" if ret else "A"
    else:
        # "Çıxan qaytarılma" = biz müştəriyə pul qaytarmışıq (debitor artır);
        # "Daxil olan qaytarılma" = təchizatçı bizə pul qaytarıb (kreditor artır)
        ret = "qayt" in n
        if n.startswith(("dax", "med", "in")):
            return "CQ" if ret else "D"
        if n.startswith(("cix", "mex", "out")):
            return "DQ" if ret else "C"
    return None


# ----------------------------------------------------------------------------
# Papkaların oxunması
# ----------------------------------------------------------------------------

# Qeyd/təyinatda bu sözlər varsa ödəniş "qaytarılan pul" sayılır (Vergi Mühasibi ilə eyni qayda)
RET_NOTE = re.compile(r"qaytar|geri|qeri|\bq/r\b|seh[vf]|sah[vf]|iade|refund|return")
POSITIONAL = {"tarix": 0, "doc": 1, "voen": 2, "name": 3, "nov": 4, "mebleg": 5, "qeyd": 6, "part": 7, "rate": 8}


def header_map(row):
    """1-ci sətir başlıqdırsa {açar: sütun nömrəsi}, yoxsa None. Sütun sırası sərbəstdir."""
    m = {}
    for c, h in enumerate(row):
        n = norm(h) if h is not None else ""
        if not n:
            continue
        if "esas" in n and "edv" in n:
            key = "part"
        elif n.startswith("edv") and "%" in n:
            key = "rate"
        elif n.startswith("tarix"):
            key = "tarix"
        elif n.startswith(("qaime", "sened")):
            key = "doc"
        elif "voen" in n or n == "tin":
            key = "voen"
        elif n.startswith("terefin") or n in ("ad", "adi"):
            key = "name"
        elif n == "nov":
            key = "nov"
        elif n.startswith("mebleg"):
            key = "mebleg"
        elif n.startswith(("qeyd", "teyinat", "izahat", "odenisin teyinat")):
            key = "qeyd"
        else:
            continue
        m.setdefault(key, c)
    return m if {"tarix", "voen", "mebleg"} <= set(m) else None   # Növ yoxdursa (bank çıxarışı) işarədən təyin olunur


def import_file(con, path, ftype):
    """Bir Excel faylını oxuyub bazaya yazır. (sətir sayı, xəta siyahısı) qaytarır."""
    errors, rows, names = [], [], {}
    try:
        wb = load_workbook(path, read_only=True, data_only=True)
    except Exception as e:
        return 0, [(path.name, 0, f"Fayl açılmadı: {e}")], False
    try:
        ws = wb["Data"] if "Data" in wb.sheetnames else wb.worksheets[0]
        try:
            ws.reset_dimensions()
        except Exception:
            pass
        idx = dict(POSITIONAL)
        for i, r in enumerate(ws.iter_rows(min_row=1, max_col=14, values_only=True), start=1):
            r = tuple(r) + (None,) * (14 - len(r))
            if all(x is None or str(x).strip() == "" for x in r):
                continue
            if i == 1:
                hm = header_map(r)
                if hm:
                    idx = hm      # sütunlar başlığa görə tapıldı (sıra fərq etmir)
                    continue
                if r[0] is not None and norm(r[0]).startswith("tarix"):
                    continue
            g = lambda key: r[idx[key]] if key in idx else None
            tarix, doc, voen, name, nov, mebleg, qeyd = (g(k) for k in
                                                         ("tarix", "doc", "voen", "name", "nov", "mebleg", "qeyd"))
            h8, h9 = g("part"), g("rate")
            d = parse_date(tarix)
            if not d:
                errors.append((path.name, i, "Tarix düzgün deyil"))
                continue
            bank = ftype == "O" and "nov" not in idx and (nov is None or str(nov).strip() == "")
            k = None if bank else kind_of(ftype, nov)
            if bank:                 # bank çıxarışı: + daxil olan, - çıxan
                try:
                    k = "D" if parse_amount(mebleg) >= 0 else "C"
                except ValueError:
                    pass
            if not k:
                errors.append((path.name, i, f"Növ tanınmadı: {nov!r}"))
                continue
            if k in ("D", "C") and RET_NOTE.search(norm(qeyd or "")):
                k = "CQ" if k == "D" else "DQ"
            v = parse_voen(voen)
            if not v:
                errors.append((path.name, i, "VÖEN boşdur"))
                continue
            try:
                amt = parse_amount(mebleg)
            except ValueError as e:
                errors.append((path.name, i, str(e)))
                continue
            if k in ("S", "A") and amt < 0:      # mənfi yazılıbsa = qaytarılma
                k, amt = k + "Q", -amt
            elif k in ("SQ", "AQ") or bank:      # qaytarılma / bank sətri həmişə müsbət saxlanır
                amt = abs(amt)
            part = parse_part(h8)
            # köhnə şablon: 8-ci sütun ƏDV % idisə
            rate_raw = h9 if h9 not in (None, "") else (
                h8 if isinstance(h8, (int, float)) and idx is POSITIONAL else None)
            if rate_raw is None:
                rate = VAT_DEFAULT if ftype == "Q" else 0
            else:
                rate = parse_vat(rate_raw)
            rows.append((str(path), i, k, d,
                         "" if doc is None else str(doc).strip(), v, amt,
                         "" if qeyd is None else str(qeyd).strip(),
                         rate, part))
            if name is not None and str(name).strip():
                names[v] = str(name).strip()
    finally:
        wb.close()

    st = path.stat()
    with con:
        con.execute("DELETE FROM txn WHERE src=?", (str(path),))
        con.executemany(
            "INSERT INTO txn(src,row_no,kind,date,doc_no,voen,amount,note,vat,part) VALUES (?,?,?,?,?,?,?,?,?,?)",
            rows)
        con.executemany(
            "INSERT INTO parties(voen,name) VALUES(?,?) "
            "ON CONFLICT(voen) DO UPDATE SET name=excluded.name", names.items())
        con.execute("INSERT OR REPLACE INTO files VALUES(?,?,?)",
                    (str(path), st.st_mtime, st.st_size))
    return len(rows), errors, True


def scan(con, cfg, progress=None):
    """Papkaları yoxlayır: yalnız yeni/dəyişmiş faylları oxuyur, silinmişləri bazadan çıxarır."""
    t0 = time.time()
    files = {}
    for ftype, key in (("Q", "qaime"), ("O", "odenis")):
        d = Path(cfg[key])
        d.mkdir(parents=True, exist_ok=True)
        for p in sorted(d.rglob("*")):
            if p.suffix.lower() in (".xlsx", ".xlsm") and not p.name.startswith("~$"):
                files[str(p)] = (ftype, p)

    known = {r[0]: (r[1], r[2]) for r in con.execute("SELECT path,mtime,size FROM files")}
    stats = {"yeni": 0, "yenilenen": 0, "kecilen": 0, "silinen": 0, "setir": 0, "xetalar": []}

    for sp in set(known) - set(files):
        with con:
            con.execute("DELETE FROM txn WHERE src=?", (sp,))
            con.execute("DELETE FROM files WHERE path=?", (sp,))
        stats["silinen"] += 1

    for n, (sp, (ftype, p)) in enumerate(files.items(), start=1):
        st = p.stat()
        if known.get(sp) == (st.st_mtime, st.st_size):
            stats["kecilen"] += 1
            continue
        if progress:
            progress(f"Oxunur ({n}/{len(files)}): {p.name}")
        cnt, errs, ok = import_file(con, p, ftype)
        stats["xetalar"].extend(errs)
        if ok:
            stats["yenilenen" if sp in known else "yeni"] += 1
            stats["setir"] += cnt

    changed = stats["yeni"] + stats["yenilenen"] + stats["silinen"]
    stats["dublikat"] = duplicates(con) if changed else 0
    stats["muddet"] = time.time() - t0
    return stats


def duplicates(con):
    """Eyni (növ, VÖEN, tarix, sənəd №, məbləğ) təkrarlanan qrupların sayı."""
    return con.execute(
        "SELECT COUNT(*) FROM (SELECT 1 FROM txn "
        "GROUP BY kind, voen, date, doc_no, amount HAVING COUNT(*)>1)").fetchone()[0]


# ----------------------------------------------------------------------------
# Hesablama
# ----------------------------------------------------------------------------

SIGN = {"S": 1, "SQ": -1, "D": -1, "DQ": 1, "A": -1, "AQ": 1, "C": 1, "CQ": -1}  # xalis: müsbət = tərəf bizə borcludur


def balances(con, as_of=None, start=None):
    """Hər tərəf üzrə: {növ: [cəm, əsas, ƏDV]} (qəpiklə). Qaytarıldı: [(voen, ad, {növ: [...]})]."""
    data = {}
    for voen, k, amt, vat, part, dte in con.execute(
            "SELECT voen, kind, amount, vat, part, date FROM txn WHERE (? IS NULL OR date <= ?)", (as_of, as_of)):
        b, v = split_amount(amt, part, vat)
        if start and dte < start:
            k = "o" + k            # başlanğıc tarixdən əvvəl = əvvəlki qalıq
        e = data.setdefault(voen, {}).setdefault(k, [0, 0, 0])
        e[0] += amt; e[1] += b; e[2] += v
    names = dict(con.execute("SELECT voen, COALESCE(NULLIF(reg_name,''), name, '') FROM parties"))
    out = [(v, names.get(v) or "", d) for v, d in data.items()]
    out.sort(key=lambda x: (x[1], x[0]))
    return out


def party_info(con):
    """{VÖEN: (ƏDV ödəyicisi, Riskli)} - portaldan gətirilmiş məlumat."""
    return {v: (a or "", b or "") for v, a, b in con.execute("SELECT voen, vat, risky FROM parties")}


def _row(v, n, d, info, plus, minus, pay, payret):
    g = lambda k, i: d.get(k, [0, 0, 0])[i]
    net = lambda pre, i: g(pre + plus, i) - g(pre + minus, i) - g(pre + pay, i) + g(pre + payret, i)
    ob, ov = net("o", 1), net("o", 2)                      # əvvəlki qalıq
    vals = []
    for k in (plus, minus, pay, payret):
        vals += [g(k, 1), g(k, 2)]          # əsas, ƏDV (dövr üzrə)
    qb = ob + vals[0] - vals[2] - vals[4] + vals[6]
    qv = ov + vals[1] - vals[3] - vals[5] + vals[7]
    vat, risky = (info or {}).get(v, ("", ""))
    return (v, n, vat, risky, ob, ov, *vals, qb, qv, qb + qv)


def debitor_rows(bal, info=None):
    """(VÖEN, ad, ƏDV ödəyicisi, riskli, əvvəlki qalıq əsas/ƏDV, satış, qaytarılma, daxil olan, qaytarılan pul
    (hər biri əsas/ƏDV), qalıq əsas, qalıq ƏDV, qalıq cəmi) - qalıq != 0."""
    rows = (_row(v, n, d, info, "S", "SQ", "D", "DQ") for v, n, d in bal)
    return [r for r in rows if r[16] != 0]


def kreditor_rows(bal, info=None):
    rows = (_row(v, n, d, info, "A", "AQ", "C", "CQ") for v, n, d in bal)
    return [r for r in rows if r[16] != 0]


def statement(con, voen, as_of=None):
    """Hərəkət (çıxarış). Hər sətir:
    (tarix, sənəd, əməliyyat,
     qaimə əsas, qaimə ƏDV, qaytarılma əsas, qaytarılma ƏDV, ödəniş əsas, ödəniş ƏDV,
     qalıq əsas, qalıq ƏDV, qeyd)
    Qalıq xalisdir: müsbət = tərəf bizə borcludur (debitor), mənfi = biz borcluyuq (kreditor)."""
    rows = con.execute(
        "SELECT date, doc_no, kind, amount, note, vat, part FROM txn "
        "WHERE voen=? AND (? IS NULL OR date<=?) ORDER BY date, id", (voen, as_of, as_of)).fetchall()
    out, qb, qv = [], 0, 0
    for date, doc, k, amt, note, vat, part in rows:
        b, v = split_amount(amt, part, vat)
        qb += SIGN[k] * b
        qv += SIGN[k] * v
        cells = [0] * 8      # qaimə, qaytarılma, ödəniş, qaytarılan pul  x (əsas, ƏDV)
        pos = {"S": 0, "A": 0, "SQ": 2, "AQ": 2, "D": 4, "C": 4, "DQ": 6, "CQ": 6}[k]
        cells[pos], cells[pos + 1] = b, v
        out.append((date, doc, KIND_LABEL[k], *cells, qb, qv, note))
    return out


def fmt_money(q):
    return f"{q / 100:,.2f}".replace(",", " ")


def iso_to_az(iso):
    return dt.date.fromisoformat(iso).strftime("%d.%m.%Y")


# ----------------------------------------------------------------------------
# Portal sorğusu (new.e-taxes.gov.az - açıq "vergi ödəyicisi məlumatı" API-si)
# ----------------------------------------------------------------------------

API_URL = "https://new.e-taxes.gov.az/api/po/authless/public/v1/authless/findTaxpayer"
_last_call = [0.0]


def norm_tin(v):
    s = re.sub(r"[^0-9]", "", str(v or ""))
    return s.zfill(10) if 0 < len(s) < 10 else s


def _post(tin, typ, sid):
    h = {"Content-Type": "application/json", "Accept": "application/json, text/plain, */*",
         "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/120 Safari/537.36",
         "Origin": "https://new.e-taxes.gov.az", "Referer": "https://new.e-taxes.gov.az/"}
    if sid:
        h["sessionId"] = sid
    req = urllib.request.Request(API_URL, json.dumps({"tin": tin, "type": typ}).encode(), h, method="POST")
    wait = 0.35 - (time.time() - _last_call[0])
    if wait > 0:
        time.sleep(wait)
    _last_call[0] = time.time()
    try:
        with urllib.request.urlopen(req, timeout=25) as r:
            return r.status, json.loads(r.read().decode("utf-8") or "null")
    except urllib.error.HTTPError as e:
        try:
            return e.code, json.loads(e.read().decode("utf-8"))
        except Exception:
            return e.code, None


def _yn(x):
    return "Bəli" if x is True else "Xeyr" if x is False else ""


def lookup_voen(tin, sid=""):
    """Portaldan VÖEN məlumatı. Qaytarır: {"ok":True,name,vat,risky,active} və ya
    {"ok":False,"err":mətn,"notfound":bool,"net":bool}."""
    order = ["legalEntity", "naturalPerson"] if tin[-1:] == "1" else ["naturalPerson", "legalEntity"]
    last = (0, None)
    try:
        for typ in order:
            for attempt in range(1, 5):
                st, body = _post(tin, typ, sid)
                if st == 429 or st >= 500:
                    time.sleep(3 * attempt)
                    continue
                break
            last = (st, body)
            tps = (body or {}).get("taxpayers") if isinstance(body, dict) else None
            if st == 200 and tps:
                tp = tps[0]
                L = tp.get("legalTaxpayerStatus") or {}
                return {"ok": True, "name": tp.get("name") or L.get("name") or "",
                        "vat": _yn(tp.get("vatPayer")), "risky": _yn(tp.get("riskyPayer")),
                        "active": _yn(tp.get("active"))}
            if isinstance(body, dict) and body.get("applicationErrorCode") == "authless-104":
                continue            # VÖEN bu növə uyğun deyil -> o biri növ
            if st == 404:
                return {"ok": False, "err": "Portalda tapılmadı", "notfound": True}
            break
    except Exception as e:
        return {"ok": False, "err": f"şəbəkə xətası: {e}", "net": True}
    msg = ""
    if isinstance(last[1], dict):
        m = last[1].get("message")
        msg = (m.get("az") or m.get("en")) if isinstance(m, dict) else (last[1].get("debug") or "")
    return {"ok": False, "err": msg or f"HTTP {last[0]}", "notfound": False}


def voens_to_check(con, force=False):
    sql = ("SELECT DISTINCT t.voen FROM txn t LEFT JOIN parties p ON p.voen=t.voen "
           + ("" if force else "WHERE p.checked IS NULL OR p.checked='' "))
    out = []
    for (v,) in con.execute(sql):
        if re.fullmatch(r"[0-9]{10}", v or ""):
            out.append(v)
    return sorted(out)


def save_info(con, voen, r):
    now = dt.date.today().isoformat()
    if r["ok"]:
        vals = (voen, r["name"], r["vat"], r["risky"], r["active"], now)
    else:   # portalda tapılmadı - təkrar sorğulanmasın
        vals = (voen, "Portalda tapılmadı", "", "", "", now)
    with con:
        con.execute(
            "INSERT INTO parties(voen,reg_name,vat,risky,active,checked) VALUES(?,?,?,?,?,?) "
            "ON CONFLICT(voen) DO UPDATE SET reg_name=excluded.reg_name, vat=excluded.vat, "
            "risky=excluded.risky, active=excluded.active, checked=excluded.checked", vals)


def check_voens(db_path, sid, todo, prog, stop=lambda: False):
    """Verilmiş VÖEN-ləri sorğulayıb bazaya yazır. prog(done, total, err, mesaj). Şəbəkə problemində dayanır."""
    con = connect(db_path)
    done = errs = netfail = 0
    msg = ""
    for tin in todo:
        if stop():
            break
        r = lookup_voen(tin, sid)
        if r["ok"] or r.get("notfound"):
            save_info(con, tin, r)
            netfail = 0
        else:
            errs += 1
            msg = r["err"]
            netfail = netfail + 1 if r.get("net") or "HTTP" in msg else 0
        done += 1
        prog(done, len(todo), errs, msg)
        if netfail >= 4:
            msg = "Portala qoşulmaq olmur (" + msg + "). Sonra yenidən cəhd edin."
            prog(done, len(todo), errs, msg)
            break
    con.close()
    return done, errs, msg


# ----------------------------------------------------------------------------
# Uzantıdan birbaşa qəbul (yalnız bu kompüterdə, 127.0.0.1)
# ----------------------------------------------------------------------------

LISTEN_PORT = 8765
SRC_EXT = "uzanti"     # uzantıdan gələn sətirlərin mənbə adı


def ingest(db_path, p):
    """Uzantının göndərdiyi qaimə sətirlərini bazaya yazır. Həmin tarix aralığındakı köhnə
    uzantı sətirləri silinib təzələri yazılır (dublikat yaranmır)."""
    company = re.sub(r"[^0-9]", "", str(p.get("company") or ""))
    if len(company) == 10:      # bir proqram bazası = bir şirkət: başqa şirkətin qaimələri qarışmasın
        cfg = load_cfg()
        if cfg.get("company") and cfg["company"] != company:
            raise ValueError(f"Bu proqram {cfg['company']} VÖEN-li şirkətə aiddir, gələn məlumat {company} şirkətinindir. "
                             "Başqa şirkət üçün proqramı ayrı qovluğa köçürüb orada işlədin.")
        if not cfg.get("company"):
            cfg["company"] = company
            save_cfg(cfg)
    d_from, d_to = parse_date(p.get("from")), parse_date(p.get("to"))
    if not d_from or not d_to:
        raise ValueError("aralıq tarixi səhvdir")
    rows, names, bad = [], {}, 0
    for i, r in enumerate(p.get("rows", []), start=1):
        r = list(r) + [None] * (8 - len(r))
        tarix, doc, voen, name, nov, mebleg, qeyd, part = r[:8]
        d, k, v = parse_date(tarix), kind_of("Q", nov) or kind_of("O", nov), parse_voen(voen)
        try:
            amt = parse_amount(mebleg)
        except ValueError:
            amt = None
        if not (d and k and v) or amt is None:
            bad += 1
            continue
        if k in ("S", "A") and amt < 0:
            k, amt = k + "Q", -amt
        elif k in ("SQ", "AQ"):
            amt = abs(amt)
        if k in ("D", "C") and RET_NOTE.search(norm(qeyd or "")):   # qaytarılan pul
            k = "CQ" if k == "D" else "DQ"
        rows.append((SRC_EXT, i, k, d, "" if doc is None else str(doc).strip(), v, amt,
                     "" if qeyd is None else str(qeyd).strip(), 0, parse_part(part)))
        if name and str(name).strip():
            names[v] = str(name).strip()
    con = connect(db_path)
    with con:
        con.execute("DELETE FROM txn WHERE src=? AND date BETWEEN ? AND ?", (SRC_EXT, d_from, d_to))
        con.executemany("INSERT INTO txn(src,row_no,kind,date,doc_no,voen,amount,note,vat,part) "
                        "VALUES (?,?,?,?,?,?,?,?,?,?)", rows)
        con.executemany("INSERT INTO parties(voen,name) VALUES(?,?) "
                        "ON CONFLICT(voen) DO UPDATE SET name=excluded.name", names.items())
    con.close()
    return len(rows), bad


def start_server(db_path, on_data=None, port=LISTEN_PORT):
    """127.0.0.1-də qəbuledici. CORS cavabı verilmir: yalnız uzantı (host icazəsi ilə) yaza bilir,
    adi veb-səhifələr yox."""
    from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

    class H(BaseHTTPRequestHandler):
        def log_message(self, *a):
            pass

        def _send(self, code, obj):
            b = json.dumps(obj, ensure_ascii=False).encode("utf-8")
            self.send_response(code)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.send_header("Content-Length", str(len(b)))
            self.end_headers()
            self.wfile.write(b)

        def do_POST(self):
            origin = self.headers.get("Origin", "")
            if self.path != "/qaime" or self.headers.get("X-DK") != "1" or \
                    (origin and not origin.startswith("chrome-extension://")):
                return self._send(403, {"ok": False, "err": "icazə yoxdur"})
            try:
                n = int(self.headers.get("Content-Length", "0"))
                if n <= 0 or n > 100_000_000:
                    raise ValueError("ölçü səhvdir")
                payload = json.loads(self.rfile.read(n).decode("utf-8"))
                cnt, bad = ingest(db_path, payload)
            except Exception as e:
                return self._send(400, {"ok": False, "err": str(e)})
            if on_data:
                on_data(cnt)
            self._send(200, {"ok": True, "rows": cnt, "skipped": bad})

    srv = ThreadingHTTPServer(("127.0.0.1", port), H)
    threading.Thread(target=srv.serve_forever, daemon=True).start()
    return srv


# ----------------------------------------------------------------------------
# Excel: şablonlar və ixrac
# ----------------------------------------------------------------------------

FONT = "Arial"
HEAD_FILL = PatternFill("solid", start_color="1F3864")
TOTAL_FILL = PatternFill("solid", start_color="D9E1F2")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
MONEY_FMT = '#,##0.00;[Red]-#,##0.00;"-"'


def _style_header(ws, ncols, row=1):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = Font(name=FONT, bold=True, color="FFFFFF")
        cell.fill = HEAD_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER
    ws.row_dimensions[row].height = 30


def make_templates(folder):
    folder = Path(folder)
    folder.mkdir(parents=True, exist_ok=True)
    specs = [
        ("sablon_qaime.xlsx", "Qaimə №", "Satış,Alış,Satış qaytarılma,Alış qaytarılma", ["Satış", "Alış"],
         ["Satış = müştəriyə kəsilən qaimə (debitor artır).",
          "Alış = təchizatçıdan alınan qaimə (kreditor artır).",
          "Satış qaytarılma = müştərinin qaytardığı mal (debitor azalır). Alış qaytarılma = təchizatçıya qaytardığımız mal (kreditor azalır).",
          "Qaytarılmada məbləği MÜSBƏT yazın (mənfi yazsanız da proqram özü qaytarılma sayır).",
          "Əsas / ƏDV sütunu: məbləğ əsas məbləğdirsə \"Əsas\", ƏDV-dirsə \"ƏDV\" seçin (bir sənəd = iki sətir).",
          "Bu sütun boşdursa məbləğ ƏDV daxil sayılır və ƏDV % (boşdursa 18) ilə ayrılır."],
         (dt.date(2026, 10, 15), "ƏDV-Q-1042", "1234567890", "Nümunə MMC", "Satış", 1180.00, "nümunə sətir")),
        ("sablon_odenis.xlsx", "Sənəd №", "Daxil olan,Çıxan,Daxil olan qaytarılma,Çıxan qaytarılma", ["Daxil olan", "Çıxan"],
         ["Daxil olan = müştəridən bizə daxil olan ödəniş (debitor azalır).",
          "Çıxan = bizdən təchizatçıya çıxan ödəniş (kreditor azalır).",
          "Əsas / ƏDV sütunu: ödənişin əsas hissəsi \"Əsas\", ƏDV hissəsi \"ƏDV\" (iki sətir). Boşdursa hamısı əsas sayılır.",
          "Daxil olan qaytarılma = təchizatçının bizə qaytardığı pul (kreditor artır). Çıxan qaytarılma = bizim müştəriyə qaytardığımız pul (debitor artır).",
          "Qeyd sütununda \"qaytarılma / geri / iadə / refund / return\" sözü varsa, ödəniş avtomatik qaytarılan pul sayılır.",
          "Bank çıxarışının sətirlərini bu şablona köçürüb papkaya atın."],
         (dt.date(2026, 10, 20), "MO-77", "1234567890", "Nümunə MMC", "Daxil olan", 500.00, "nümunə sətir")),
    ]
    for fname, docname, validation, types, notes, example in specs:
        wb = Workbook()
        ws = wb.active
        ws.title = "Data"
        headers = ["Tarix", docname, "Əsas / ƏDV", "VÖEN", "Tərəfin adı", "Növ", "Məbləğ", "Qeyd", "ƏDV %"]
        ws.append(headers)
        _style_header(ws, len(headers))
        for col, w in zip("ABCDEFGHI", (13, 16, 13, 15, 34, 20, 16, 30, 9)):
            ws.column_dimensions[col].width = w
        for r in range(2, 2002):
            ws.cell(row=r, column=1).number_format = "DD.MM.YYYY"
            ws.cell(row=r, column=4).number_format = "@"
            ws.cell(row=r, column=7).number_format = "#,##0.00"
            for c in range(1, 10):
                ws.cell(row=r, column=c).font = Font(name=FONT)
        dv = DataValidation(type="list", formula1=f'"{validation}"', allow_blank=True)
        ws.add_data_validation(dv)
        dv.add("F2:F5000")
        dv2 = DataValidation(type="list", formula1='"Əsas,ƏDV"', allow_blank=True)
        ws.add_data_validation(dv2)
        dv2.add("C2:C5000")
        ws.freeze_panes = "A2"

        info = wb.create_sheet("Təlimat")
        info.column_dimensions["A"].width = 100
        lines = ["Bu şablonu doldurub müvafiq papkaya qoyun, sonra proqramda \"Papkaları yenilə\" basın.",
                 "Məlumat yalnız \"Data\" vərəqindən oxunur. Sütunlar başlıq adına görə tapılır, sırası sərbəstdir (başlıq sətrini silməyin).",
                 "Məcburi sütunlar: Tarix, VÖEN, Növ, Məbləğ. Tərəfin adı, ƏDV ödəyicisi və riskli statusu VÖEN üzrə portaldan avtomatik gətirilir (ad xanasını boş qoya bilərsiniz).",
                 ""] + notes + ["", "Nümunə sətir (bunu Data vərəqinə köçürməyin):"]
        for ln in lines:
            info.append([ln])
        info.append(["   |   ".join([example[0].strftime("%d.%m.%Y"), example[1], "Əsas", example[2],
                                     example[3], example[4], f"{example[5]:.2f}", example[6]])])
        for row in info.iter_rows():
            for cell in row:
                cell.font = Font(name=FONT, bold=(cell.row == 1))
        wb.save(folder / fname)
    wb = Workbook()
    ws = wb.active
    ws.title = "Data"
    ws.append(["VÖEN", "Tarix", "Məbləğ (+ daxil olan, − çıxan)", "Təyinat"])
    _style_header(ws, 4)
    for col, w in zip("ABCD", (15, 13, 28, 60)):
        ws.column_dimensions[col].width = w
    ws.append(["1234567890", dt.date(2026, 10, 20), 500.00, "nümunə: müştəridən daxilolma"])
    ws.append(["1234567890", dt.date(2026, 10, 25), -50.00, "nümunə: pulun qaytarılması (təyinatda 'qaytarılma' sözü)"])
    ws.append([])
    ws.append(["Bank çıxarışı: hansı bankdan olursa olsun 4 sütuna köçürüb odenisler papkasına atın. Növ sütunu lazım deyil."])
    ws.append(["Məbləğ + = bizə daxil olan, − = bizdən çıxan. Təyinatda qaytarılma / geri / iadə / səhv / refund / return varsa qaytarılan pul sayılır."])
    ws.append(["Nümunə sətirləri silin."])
    wb.save(folder / "sablon_bank.xlsx")
    return [folder / s[0] for s in specs] + [folder / "sablon_bank.xlsx"]


def _write_table(ws, headers, rows, money_cols, formula_cols=None, text_cols=(), total=True):
    """rows - siyahı; formula_cols = {sütun_nömrəsi: lambda sətir_nömrəsi -> formula}"""
    formula_cols = formula_cols or {}
    ws.append(headers)
    _style_header(ws, len(headers))
    for r_i, row in enumerate(rows, start=2):
        vals = list(row)
        for c, fn in formula_cols.items():
            vals[c - 1] = fn(r_i)
        ws.append(vals)
        for c in range(1, len(headers) + 1):
            cell = ws.cell(row=r_i, column=c)
            cell.font = Font(name=FONT)
            cell.border = BORDER
            if c in money_cols:
                cell.number_format = MONEY_FMT
            if c in text_cols:
                cell.number_format = "@"
    last = len(rows) + 1
    if total and rows:
        tr = last + 1
        ws.cell(row=tr, column=1, value="YEKUN")
        for c in money_cols:
            L = get_column_letter(c)
            ws.cell(row=tr, column=c, value=f"=SUM({L}2:{L}{last})")
        for c in range(1, len(headers) + 1):
            cell = ws.cell(row=tr, column=c)
            cell.font = Font(name=FONT, bold=True)
            cell.fill = TOTAL_FILL
            cell.border = BORDER
            if c in money_cols:
                cell.number_format = MONEY_FMT
    ws.freeze_panes = "A2"
    if rows:
        ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{last}"


def export_xlsx(path, con, as_of=None, start=None):
    bal = balances(con, as_of, start)
    wb = Workbook()
    ws = wb.active
    ws.title = "Debitor"
    f = lambda r: tuple(x / 100 if isinstance(x, int) else x for x in r)
    info = party_info(con)
    hd = ["VÖEN", "Tərəfin adı", "ƏDV ödəyicisi", "Riskli"]
    mc = tuple(range(5, 18))
    fm = {15: lambda r: f"=E{r}+G{r}-I{r}-K{r}+M{r}", 16: lambda r: f"=F{r}+H{r}-J{r}-L{r}+N{r}", 17: lambda r: f"=O{r}+P{r}"}
    ws = wb["Debitor"]
    _write_table(ws, hd + ["Əvvəlki qalıq əsas", "Əvvəlki qalıq ƏDV", "Satış əsas", "Satış ƏDV", "Satış qaytarılma əsas", "Satış qaytarılma ƏDV",
                           "Ödəniş əsas", "Ödəniş ƏDV", "Qaytarılan pul əsas", "Qaytarılan pul ƏDV",
                           "Qalıq əsas", "Qalıq ƏDV", "Qalıq cəmi"],
                 [f(r) for r in debitor_rows(bal, info)], money_cols=mc, formula_cols=fm, text_cols=(1,))
    ws2 = wb.create_sheet("Kreditor")
    _write_table(ws2, hd + ["Əvvəlki qalıq əsas", "Əvvəlki qalıq ƏDV", "Alış əsas", "Alış ƏDV", "Alış qaytarılma əsas", "Alış qaytarılma ƏDV",
                            "Çıxan ödəniş əsas", "Çıxan ödəniş ƏDV", "Qaytarılan pul əsas", "Qaytarılan pul ƏDV",
                            "Qalıq əsas", "Qalıq ƏDV", "Qalıq cəmi"],
                 [f(r) for r in kreditor_rows(bal, info)], money_cols=mc, formula_cols=fm, text_cols=(1,))
    for w in (ws, ws2):
        for col, width in zip("ABCDEFGHIJKLMNOPQ", (15, 40, 13, 9, 16, 16, 15, 15, 17, 17, 15, 15, 16, 16, 15, 15, 16)):
            w.column_dimensions[col].width = width
    info = wb.create_sheet("Məlumat")
    info.append(["Hesabat aralığı", (iso_to_az(start) if start else "əvvəldən") + " - " + (iso_to_az(as_of) if as_of else "bütün əməliyyatlar")])
    info.append(["Yaradılma vaxtı", dt.datetime.now().strftime("%d.%m.%Y %H:%M")])
    info.append(["Qeyd", "Qalıq = qaimələr − qaytarılmalar − ödənişlər. Mənfi qalıq: qarşı tərəfə artıq ödəniş edilib."])
    info.column_dimensions["A"].width = 18
    info.column_dimensions["B"].width = 80
    for row in info.iter_rows():
        for cell in row:
            cell.font = Font(name=FONT, bold=(cell.column == 1))
    wb.save(path)


STMT_HEADERS = ["Tarix", "Sənəd №", "Əməliyyat", "E-qaimə əsas", "E-qaimə ƏDV",
                "Qaytarılma e-qaimə əsas", "Qaytarılma e-qaimə ƏDV", "Ödəniş əsas", "Ödəniş ƏDV",
                "Qaytarılan pul əsas", "Qaytarılan pul ƏDV", "Qalıq əsas", "Qalıq ƏDV", "Qeyd"]


def export_statement(path, con, voen, name, as_of=None):
    rows = statement(con, voen, as_of)
    wb = Workbook()
    ws = wb.active
    ws.title = "Çıxarış"
    data = [(iso_to_az(r[0]), r[1], r[2], *[x / 100 for x in r[3:13]], r[13]) for r in rows]
    _write_table(ws, STMT_HEADERS, data, money_cols=tuple(range(4, 14)), total=False)
    n = len(data)
    if n:   # yekun sətri: yalnız hərəkət sütunları (qalıq sütunları yekunlanmır)
        tr = n + 2
        ws.cell(row=tr, column=1, value="YEKUN")
        for c in range(4, 12):
            L = get_column_letter(c)
            ws.cell(row=tr, column=c, value=f"=SUM({L}4:{L}{n + 3})")  # 2 sətir yuxarıdan əlavə olunacaq.number_format = MONEY_FMT
        for c in range(1, len(STMT_HEADERS) + 1):
            cell = ws.cell(row=tr, column=c)
            cell.font = Font(name=FONT, bold=True)
            cell.fill = TOTAL_FILL
            cell.border = BORDER
    for col, width in zip("ABCDEFGHIJKLMN", (13, 18, 24, 15, 15, 17, 17, 15, 15, 15, 15, 15, 15, 30)):
        ws.column_dimensions[col].width = width
    ws.insert_rows(1, amount=2)
    ws["A1"] = f"{name}  (VÖEN: {voen})"
    ws["A1"].font = Font(name=FONT, bold=True, size=12)
    ws["A2"] = "Tarixə qədər: " + (iso_to_az(as_of) if as_of else "bütün əməliyyatlar")
    ws["A2"].font = Font(name=FONT, italic=True)
    ws.freeze_panes = "A4"
    wb.save(path)


# ----------------------------------------------------------------------------
# Pəncərəli proqram (Tkinter)
# ----------------------------------------------------------------------------

def run_gui():
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox

    cfg = load_cfg()
    con = connect(cfg["db"])
    root = tk.Tk()
    root.title("Debitor-Kreditor")
    root.geometry("1700x680")

    state = {"bal": [], "as_of": None}
    status = tk.StringVar(value="Hazır")
    asof_var = tk.StringVar()
    from_var = tk.StringVar()
    search_var = tk.StringVar()

    class Table:
        def __init__(self, parent, headers, money_from=2):
            self.mf = money_from
            frame = ttk.Frame(parent)
            self.frame = frame
            self.tree = ttk.Treeview(frame, columns=list(range(len(headers))), show="headings")
            sb = ttk.Scrollbar(frame, orient="vertical", command=self.tree.yview)
            self.tree.configure(yscrollcommand=sb.set)
            self.tree.pack(side="left", fill="both", expand=True)
            sb.pack(side="right", fill="y")
            self.rows, self.rev = [], False
            for i, h in enumerate(headers):
                self.tree.heading(i, text=h, command=lambda i=i: self.sort(i))
                self.tree.column(i, width=(105 if i >= money_from else 90) if i != 1 else 260,
                                 anchor="e" if i >= money_from else "w")
            self.total_label = None
            self.tree.tag_configure("risky", foreground="#c53030")

        def fill(self, rows):
            self.rows = rows
            self.show(rows)

        def show(self, rows):
            self.tree.delete(*self.tree.get_children())
            for r in rows:
                self.tree.insert("", "end", tags=("risky",) if len(r) > 3 and r[3] == "Bəli" else (),
                                 values=[fmt_money(x) if i >= self.mf else x for i, x in enumerate(r)])

        def sort(self, i):
            self.rev = not self.rev
            self.rows.sort(key=lambda r: r[i], reverse=self.rev)
            apply_filter()

    nb = ttk.Notebook(root)
    deb_t = Table(nb, ["VÖEN", "Tərəfin adı", "ƏDV ödəyicisi", "Riskli", "Əvvəlki qalıq əsas",
                       "Əvvəlki qalıq ƏDV", "Satış əsas", "Satış ƏDV", "Qaytarılma əsas", "Qaytarılma ƏDV",
                       "Daxil olan əsas", "Daxil olan ƏDV", "Qaytarılan pul əsas", "Qaytarılan pul ƏDV",
                       "Qalıq əsas", "Qalıq ƏDV", "Qalıq cəmi"], money_from=4)
    kred_t = Table(nb, ["VÖEN", "Tərəfin adı", "ƏDV ödəyicisi", "Riskli", "Əvvəlki qalıq əsas",
                        "Əvvəlki qalıq ƏDV", "Alış əsas", "Alış ƏDV", "Qaytarılma əsas", "Qaytarılma ƏDV",
                        "Çıxan əsas", "Çıxan ƏDV", "Qaytarılan pul əsas", "Qaytarılan pul ƏDV",
                        "Qalıq əsas", "Qalıq ƏDV", "Qalıq cəmi"], money_from=4)
    nb.add(deb_t.frame, text="  Debitor  ")
    nb.add(kred_t.frame, text="  Kreditor  ")
    totals = tk.StringVar()

    def apply_filter(*_):
        q = norm(search_var.get())
        for t in (deb_t, kred_t):
            rows = [r for r in t.rows if not q or q in norm(r[0]) or q in norm(r[1])]
            t.show(rows)
        d = sum(r[16] for r in deb_t.rows)
        k = sum(r[16] for r in kred_t.rows)
        totals.set(f"Debitor yekun: {fmt_money(d)}      Kreditor yekun: {fmt_money(k)}      "
                   f"Xalis: {fmt_money(d - k)}")

    def refresh():
        txt = asof_var.get().strip()
        as_of = parse_date(txt) if txt else None
        if txt and not as_of:
            messagebox.showerror("Tarix", "Tarixi gün.ay.il formatında yazın, məs. 31.12.2026")
            return
        ftxt = from_var.get().strip()
        start = parse_date(ftxt) if ftxt else None
        if ftxt and not start:
            messagebox.showerror("Tarix", "Başlanğıc tarixi gün.ay.il formatında yazın, məs. 01.01.2026")
            return
        state["as_of"] = as_of
        state["start"] = start
        bal = balances(con, as_of, start)
        state["bal"] = bal
        info = party_info(con)
        deb_t.fill(debitor_rows(bal, info))
        kred_t.fill(kreditor_rows(bal, info))
        apply_filter()

    def do_scan():
        root.config(cursor="watch")
        try:
            def prog(m):
                status.set(m)
                root.update_idletasks()
            s = scan(con, load_cfg(), prog)
        finally:
            root.config(cursor="")
        refresh()
        msg = (f"Yeni fayl: {s['yeni']}, yenilənən: {s['yenilenen']}, dəyişməyən: {s['kecilen']}, "
               f"silinən: {s['silinen']}, oxunan sətir: {s['setir']}  ({s['muddet']:.1f} san)")
        status.set(msg)
        start_lookup()
        warn = []
        if s["dublikat"]:
            warn.append(f"Diqqət: {s['dublikat']} qrup təkrarlanan əməliyyat var "
                        "(eyni növ, VÖEN, tarix, sənəd № və məbləğ). Eyni qaimə iki faylda ola bilər.")
        if s["xetalar"]:
            lines = [f"{f}, sətir {r}: {m}" for f, r, m in s["xetalar"][:25]]
            more = len(s["xetalar"]) - len(lines)
            warn.append(f"{len(s['xetalar'])} sətir oxunmadı:\n" + "\n".join(lines) +
                        (f"\n... və daha {more}" if more > 0 else ""))
        if warn:
            messagebox.showwarning("Yeniləmə nəticəsi", "\n\n".join(warn))

    def start_lookup(force=False):
        if state.get("lk"):
            return
        todo = voens_to_check(con, force)
        if not todo:
            if force:
                status.set("Yoxlanacaq VÖEN yoxdur.")
            return
        q = {"done": 0, "total": len(todo), "err": 0, "msg": "", "fin": False}
        state["lk"] = q
        sid = load_cfg().get("sessionId", "")

        def prog(d, t, e, m):
            q.update(done=d, err=e, msg=m)

        def work():
            try:
                check_voens(cfg["db"], sid, todo, prog)
            except Exception as ex:
                q["msg"] = str(ex)
            q["fin"] = True

        threading.Thread(target=work, daemon=True).start()

        def poll():
            txt = f"VÖEN portaldan yoxlanılır: {q['done']}/{q['total']}"
            if q["err"]:
                txt += f" (xəta: {q['err']}: {q['msg']})"
            status.set(txt)
            if q["fin"]:
                state["lk"] = None
                refresh()
                status.set(txt.replace("yoxlanılır", "yoxlandı"))
            else:
                if q["done"] % 5 == 0:
                    refresh()
                root.after(700, poll)
        poll()

    def do_export():
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")],
            initialfile=f"debitor_kreditor_{dt.date.today():%Y-%m-%d}.xlsx")
        if path:
            export_xlsx(path, con, state["as_of"], state.get("start"))
            status.set(f"İxrac edildi: {path}")

    def do_templates():
        files = make_templates(BASE / "sablonlar")
        messagebox.showinfo("Şablonlar", "Yaradıldı:\n" + "\n".join(str(f) for f in files))

    def pick_folder(key, title):
        d = filedialog.askdirectory(title=title, initialdir=load_cfg()[key])
        if d:
            c = load_cfg()
            c[key] = d
            save_cfg(c)
            do_scan()

    def open_statement(_event=None):
        tree = deb_t.tree if nb.index(nb.select()) == 0 else kred_t.tree
        sel = tree.selection()
        if not sel:
            return
        vals = tree.item(sel[0])["values"]
        voen, name = str(vals[0]), str(vals[1])
        rows = statement(con, voen, state["as_of"])
        win = tk.Toplevel(root)
        win.title(f"Çıxarış - {name} ({voen})")
        win.geometry("1000x500")
        t = Table(win, STMT_HEADERS, money_from=3)
        t.frame.pack(fill="both", expand=True)
        win.geometry("1750x500")
        for i in range(3, 13):
            t.tree.column(i, width=105)
        t.tree.column(1, width=110)
        t.tree.column(2, width=170, anchor="w")
        t.tree.column(13, width=180, anchor="w")
        z = lambda x: fmt_money(x) if x else ""
        for r in rows:
            t.tree.insert("", "end", values=[iso_to_az(r[0]), r[1], r[2]] + [z(x) for x in r[3:11]] +
                                            [fmt_money(r[11]), fmt_money(r[12]), r[13]])

        def exp():
            p = filedialog.asksaveasfilename(parent=win, defaultextension=".xlsx",
                                             filetypes=[("Excel", "*.xlsx")],
                                             initialfile=f"cixaris_{voen}.xlsx")
            if p:
                export_statement(p, con, voen, name, state["as_of"])

        ttk.Button(win, text="Excelə ixrac", command=exp).pack(pady=6)

    bar = ttk.Frame(root, padding=6)
    bar.pack(fill="x")
    ttk.Button(bar, text="Papkaları yenilə", command=do_scan).pack(side="left", padx=2)
    ttk.Button(bar, text="Excelə ixrac", command=do_export).pack(side="left", padx=2)
    ttk.Button(bar, text="Şablonları yarat", command=do_templates).pack(side="left", padx=2)
    ttk.Button(bar, text="VÖEN yoxla (portal)",
               command=lambda: start_lookup(True)).pack(side="left", padx=2)
    ttk.Button(bar, text="Qaimə papkası…",
               command=lambda: pick_folder("qaime", "Qaimələr papkası")).pack(side="left", padx=2)
    ttk.Button(bar, text="Ödəniş papkası…",
               command=lambda: pick_folder("odenis", "Ödənişlər / çıxarışlar papkası")).pack(side="left", padx=2)
    ttk.Label(bar, text="   Tarixdən:").pack(side="left")
    e0 = ttk.Entry(bar, textvariable=from_var, width=11)
    e0.pack(side="left")
    e0.bind("<Return>", lambda _e: refresh())
    ttk.Label(bar, text=" Tarixədək:").pack(side="left")
    e = ttk.Entry(bar, textvariable=asof_var, width=11)
    e.pack(side="left")
    e.bind("<Return>", lambda _e: refresh())
    ttk.Button(bar, text="Göstər", command=refresh).pack(side="left", padx=2)
    ttk.Label(bar, text="   Axtar:").pack(side="left")
    ttk.Entry(bar, textvariable=search_var, width=22).pack(side="left")
    search_var.trace_add("write", apply_filter)

    nb.pack(fill="both", expand=True, padx=6)
    ttk.Label(root, textvariable=totals, font=("Arial", 10, "bold")).pack(anchor="w", padx=8, pady=2)
    ttk.Label(root, textvariable=status, relief="sunken", anchor="w").pack(fill="x", side="bottom")
    deb_t.tree.bind("<Double-1>", open_statement)
    kred_t.tree.bind("<Double-1>", open_statement)

    def on_data(cnt):
        state["dirty"] = cnt

    try:
        start_server(cfg["db"], on_data)
    except OSError as ex:
        status.set(f"Uzantıdan qəbul işləmir (port {LISTEN_PORT} məşğuldur: {ex}). Proqramın ikinci nüsxəsi açıq ola bilər.")

    def check_incoming():
        if state.get("dirty") is not None:
            cnt = state.pop("dirty")
            refresh()
            status.set(f"Uzantıdan {cnt} sətir alındı.")
            start_lookup()
        root.after(1000, check_incoming)

    root.after(1000, check_incoming)
    root.after(100, do_scan)
    root.mainloop()



# ============================================================================
# ERP VEB İNTERFEYSİ (yerli server, yalnız 127.0.0.1)
# Modullar: Debitor / Kreditor, Yaşlandırma, Risk və limit, Mənfəət-zərər, Parametrlər, Fayllar
# ============================================================================
PORT = 8765
BUCKETS = [(0, 30), (31, 60), (61, 90), (91, 180), (181, 10 ** 6)]
BUCKET_LABELS = ["0–30 gün", "31–60", "61–90", "91–180", "180+"]
DEFAULT_SETTINGS = {"credit_default": 0, "credit_overdue_days": 60, "credit_limits": "",
                    "allow": [0, 0, 0, 0, 0]}


# ----------------------------------------------------------------------------
# Hesablamalar
# ----------------------------------------------------------------------------

def get_settings(dk):
    s = dict(DEFAULT_SETTINGS)
    s.update(dk.load_cfg().get("erp", {}))
    return s


def parse_limits(text):
    out = {}
    for line in str(text or "").splitlines():
        q = [x.strip() for x in line.replace("\t", ";").split(";")]
        if len(q) >= 2 and q[0]:
            try:
                out[q[0]] = float(q[1].replace(",", "."))
            except ValueError:
                pass
    return out


def aging(dk, con, side, as_of):
    """Açıq borcun yaşlandırması: qalıq ən son borc sənədlərinə (sondan əvvələ) bölünür.
    side: 'deb' (S, DQ artırır; SQ, D azaldır) / 'kred' (A, CQ artırır; AQ, C azaldır).
    Qaytarır: {voen: {"name", "buckets":[5], "total"}} (qəpiklə, ƏDV daxil)."""
    plus, minus = (("S", "DQ"), ("SQ", "D")) if side == "deb" else (("A", "CQ"), ("AQ", "C"))
    docs, bal = {}, {}
    q = "SELECT voen, kind, date, amount FROM txn WHERE date <= ? ORDER BY date"
    for v, k, d, a in con.execute(q, (as_of,)):
        if k in plus:
            docs.setdefault(v, []).append([d, a])
            bal[v] = bal.get(v, 0) + a
        elif k in minus:
            bal[v] = bal.get(v, 0) - a
    names = dict(con.execute("SELECT voen, COALESCE(NULLIF(reg_name,''), name, '') FROM parties"))
    today = dt.date.fromisoformat(as_of)
    out = {}
    for v, b in bal.items():
        if b <= 0:
            continue
        rem, bk = b, [0] * 5
        for d, a in reversed(docs.get(v, [])):
            if rem <= 0:
                break
            take = min(rem, a)
            age = (today - dt.date.fromisoformat(d)).days
            i = next(i for i, (lo, hi) in enumerate(BUCKETS) if lo <= max(age, 0) <= hi)
            bk[i] += take
            rem -= take
        if rem > 0:
            bk[4] += rem
        out[v] = {"voen": v, "name": names.get(v, ""), "buckets": bk, "total": b}
    return out


def risk_rows(dk, con, as_of, set_):
    ag = aging(dk, con, "deb", as_of)
    lim = parse_limits(set_["credit_limits"])
    dflt = float(set_["credit_default"] or 0) * 100
    od = int(set_["credit_overdue_days"] or 60)
    rows = []
    for v, r in ag.items():
        debt = r["total"]
        L = lim.get(v, None)
        L = L * 100 if L is not None else dflt
        over = sum(r["buckets"][k] for k, lo in enumerate([31, 61, 91, 181], start=1) if lo > od)
        use = debt / L * 100 if L > 0 else None
        exceed = L > 0 and debt > L
        st = "Normal"
        if exceed and over > 0:
            st = "Limit aşılıb + gecikmiş"
        elif exceed:
            st = "Limit aşılıb"
        elif over > 0:
            st = "Gecikmiş borc"
        elif use is not None and use >= 80:
            st = "Limitə yaxın (80%+)"
        rows.append({"voen": v, "name": r["name"], "debt": debt, "limit": L if L > 0 else None,
                     "free": (L - debt) if L > 0 else None, "use": round(use, 1) if use is not None else None,
                     "over": over, "status": st,
                     "_o": (2 if exceed else 0) + (2 if over > 0 else 0) + (1 if use is not None and use >= 80 else 0)})
    rows.sort(key=lambda x: (-x["_o"], -x["debt"]))
    return rows


# ----------------------------------------------------------------------------
# Server
# ----------------------------------------------------------------------------

class State:
    rev = 0                 # məlumat dəyişdikcə artır (səhifə özünü yeniləsin)
    lookup = None           # {"done","total","err","msg","fin"}
    scan_msg = ""
    last_incoming = ""
    last_ping = 0.0


def serve(dk, open_browser=True):
    from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
    cfg = dk.load_cfg()
    db = cfg["db"]
    (dk.BASE / "maza_data").mkdir(exist_ok=True)
    token = secrets.token_hex(16)
    st = State()
    st.pl = {"ctx": None, "data": None, "file": "", "mtime": 0, "msg": "", "busy": False}
    lock = threading.Lock()

    def bump():
        st.rev += 1

    def do_scan():
        con = dk.connect(db)
        s = dk.scan(con, dk.load_cfg())
        con.close()
        st.scan_msg = (f"Yeni fayl: {s['yeni']}, yenilənən: {s['yenilenen']}, dəyişməyən: {s['kecilen']}, "
                       f"silinən: {s['silinen']}, oxunan sətir: {s['setir']}"
                       + (f" | XƏTA: {len(s['xetalar'])} sətir" if s["xetalar"] else "")
                       + (f" | dublikat qrup: {s['dublikat']}" if s["dublikat"] else ""))
        st.errors = s["xetalar"][:30]
        bump()

    def start_lookup(force=False):
        if st.lookup and not st.lookup["fin"]:
            return
        con = dk.connect(db)
        todo = dk.voens_to_check(con, force)
        con.close()
        if not todo:
            return
        q = {"done": 0, "total": len(todo), "err": 0, "msg": "", "fin": False}
        st.lookup = q
        sid = dk.load_cfg().get("sessionId", "")

        def work():
            try:
                dk.check_voens(db, sid, todo, lambda d, t, e, m: q.update(done=d, err=e, msg=m))
            except Exception as ex:
                q["msg"] = str(ex)
            q["fin"] = True
            bump()
        threading.Thread(target=work, daemon=True).start()

    def pl_load(force=False):
        path = dk.load_cfg().get("pl_file", "")
        p = st.pl
        if not path or p["busy"]:
            return
        try:
            mt = os.stat(path).st_mtime
        except OSError:
            p["msg"] = "Fayl tapılmadı: " + path
            return
        if not force and p["file"] == path and p["mtime"] == mt and p["ctx"]:
            return
        p["busy"], p["msg"] = True, "Excel oxunur (bir neçə saniyə)…"

        def work():
            try:
                data = pl_load_workbook(path)
                ctx = pl_build(data)
                rej = data["rejected"]
                p.update(ctx=ctx, data=data, file=path, mtime=mt,
                         msg="Oxundu" + (f" | formatı uyğun gəlməyən vərəqlər: {', '.join(rej)}" if rej else ""))
            except Exception as ex:
                p["msg"] = f"Xəta: {ex}"
            p["busy"] = False
            bump()
        threading.Thread(target=work, daemon=True).start()

    def on_data(cnt):
        st.last_incoming = f"Uzantıdan {cnt} sətir alındı ({dt.datetime.now():%H:%M:%S})"
        bump()
        start_lookup()

    class H(BaseHTTPRequestHandler):
        def log_message(self, *a):
            pass

        def _send(self, code, body, ctype="application/json; charset=utf-8", extra=None):
            if not isinstance(body, (bytes, bytearray)):
                body = json.dumps(body, ensure_ascii=False).encode("utf-8")
            self.send_response(code)
            self.send_header("Content-Type", ctype)
            self.send_header("Content-Length", str(len(body)))
            for k, v in (extra or {}).items():
                self.send_header(k, v)
            self.end_headers()
            self.wfile.write(body)

        def _auth(self):
            return self.headers.get("X-UI") == token

        def do_GET(self):
            u = urlparse(self.path)
            if u.path == "/":
                return self._send(200, PAGE.replace("__TOKEN__", token).encode("utf-8"), "text/html; charset=utf-8")
            if not u.path.startswith("/api/") or not self._auth():
                return self._send(403, {"err": "icazə yoxdur"})
            q = {k: v[0] for k, v in parse_qs(u.query).items()}
            try:
                self._send(200, self.api_get(u.path, q))
            except BinaryResult as b:
                self._send(200, b.data, b.ctype, {"Content-Disposition": f'attachment; filename="{b.name}"'})
            except Exception as e:
                self._send(500, {"err": str(e)})

        def do_POST(self):
            u = urlparse(self.path)
            n = int(self.headers.get("Content-Length", "0") or 0)
            raw = self.rfile.read(n) if 0 < n <= 100_000_000 else b""
            if u.path == "/qaime":      # uzantıdan qəbul (əvvəlki qayda)
                origin = self.headers.get("Origin", "")
                if self.headers.get("X-DK") != "1" or (origin and not origin.startswith("chrome-extension://")):
                    return self._send(403, {"ok": False, "err": "icazə yoxdur"})
                try:
                    cnt, bad = dk.ingest(db, json.loads(raw.decode("utf-8")))
                except Exception as e:
                    return self._send(400, {"ok": False, "err": str(e)})
                on_data(cnt)
                return self._send(200, {"ok": True, "rows": cnt, "skipped": bad})
            if not self._auth():
                return self._send(403, {"err": "icazə yoxdur"})
            try:
                body = json.loads(raw.decode("utf-8")) if raw else {}
                self._send(200, self.api_post(u.path, body))
            except Exception as e:
                self._send(500, {"err": str(e)})

        # ---- GET ----
        def api_get(self, path, q):
            con = dk.connect(db)
            try:
                as_of = dk.parse_date(q.get("to")) if q.get("to") else None
                start = dk.parse_date(q.get("from")) if q.get("from") else None
                if path == "/api/status":
                    lk = st.lookup
                    n = con.execute("SELECT COUNT(*) FROM txn").fetchone()[0]
                    cfg2 = dk.load_cfg()
                    return {"rev": st.rev, "rows": n, "scan": st.scan_msg, "incoming": st.last_incoming,
                            "lookup": lk, "errors": getattr(st, "errors", []),
                            "company": cfg2.get("company", ""), "qaime": cfg2["qaime"], "odenis": cfg2["odenis"],
                            "db": cfg2["db"]}
                if path == "/api/balances":
                    bal = dk.balances(con, as_of, start)
                    info = dk.party_info(con)
                    rows = (dk.debitor_rows if q.get("side") == "deb" else dk.kreditor_rows)(bal, info)
                    return {"rows": rows}
                if path == "/api/statement":
                    rows = dk.statement(con, q["voen"], as_of)
                    name = con.execute("SELECT COALESCE(NULLIF(reg_name,''), name, '') FROM parties WHERE voen=?",
                                       (q["voen"],)).fetchone()
                    return {"voen": q["voen"], "name": name[0] if name else "", "rows": rows}
                if path in ("/api/aging", "/api/risk"):
                    d = as_of or dt.date.today().isoformat()
                    s = get_settings(dk)
                    if path == "/api/risk":
                        return {"rows": risk_rows(dk, con, d, s), "settings": s}
                    out = {}
                    for side in ("deb", "kred"):
                        rows = sorted(aging(dk, con, side, d).values(), key=lambda r: -r["total"])
                        for r in rows:
                            r["reserve"] = sum(b * (float(s["allow"][i]) / 100) for i, b in enumerate(r["buckets"]))
                        out[side] = rows
                    return {"as_of": d, "labels": BUCKET_LABELS, "allow": s["allow"], **out}
                if path == "/api/settings":
                    return get_settings(dk)
                if path in ("/api/pl", "/api/pl/export"):
                    p = st.pl
                    pl_load()
                    if path == "/api/pl/export":
                        if not p["ctx"]:
                            raise ValueError("Əvvəlcə Excel iş kitabını oxudun")
                        yr = int(q["year"]) if q.get("year") else pl_default_year(p["ctx"])
                        tmp = tempfile.mktemp(suffix=".xlsx")
                        pl_export(tmp, p["ctx"], yr)
                        data = open(tmp, "rb").read()
                        raise BinaryResult(data, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                           f"menfeet_zerer_{yr}.xlsx")
                    out = {"busy": p["busy"], "msg": p["msg"], "file": dk.load_cfg().get("pl_file", ""),
                           "ready": bool(p["ctx"])}
                    if p["ctx"]:
                        ctx = p["ctx"]
                        yr = int(q["year"]) if q.get("year") else pl_default_year(ctx)
                        by = pl_compute(ctx, yr)
                        rec = pl_reconcile(ctx)
                        out.update(years=ctx["years"], year=yr, months=PL_MONTHS,
                                   rows=[{"label": l["label"], "kind": l["kind"],
                                          "m": [round(x * 100) for x in by[l["r"]]["m"]],
                                          "cum": round(by[l["r"]]["cum"] * 100)} for l in PL_LINES],
                                   rec=(None if not rec else {
                                       "year": rec["year"], "checked": rec["checked"], "n": len(rec["diffs"]),
                                       "first": [[d[1], d[2], round(d[3], 2), round(d[4], 2)] for d in rec["diffs"][:12]]}))
                    return out
                if path == "/api/export":
                    tmp = tempfile.mktemp(suffix=".xlsx")
                    dk.export_xlsx(tmp, con, as_of, start)
                    data = open(tmp, "rb").read()
                    raise BinaryResult(data, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                       f"debitor_kreditor_{dt.date.today():%Y-%m-%d}.xlsx")
                if path.startswith("/api/maza/"):
                    r = maza_get(db, path, q)
                    if isinstance(r, tuple) and r and r[0] == "bin":
                        raise BinaryResult(r[1], r[2], r[3])
                    return r
                raise ValueError("naməlum yol")
            finally:
                con.close()

        # ---- POST ----
        def api_post(self, path, body):
            if path == "/api/scan":
                do_scan()
                start_lookup()
                return {"ok": True, "msg": st.scan_msg}
            if path == "/api/lookup":
                start_lookup(bool(body.get("force")))
                return {"ok": True}
            if path == "/api/settings":
                cfg2 = dk.load_cfg()
                s = get_settings(dk)
                for k in ("credit_default", "credit_overdue_days", "credit_limits", "allow"):
                    if k in body:
                        s[k] = body[k]
                cfg2["erp"] = s
                dk.save_cfg(cfg2)
                bump()
                return {"ok": True}
            if path == "/api/folders":
                cfg2 = dk.load_cfg()
                for k in ("qaime", "odenis"):
                    if body.get(k):
                        cfg2[k] = body[k]
                dk.save_cfg(cfg2)
                do_scan()
                return {"ok": True}
            if path == "/api/templates":
                files = dk.make_templates(dk.BASE / "sablonlar")
                return {"ok": True, "files": [str(f) for f in files]}
            if path == "/api/pl/file":
                cfg2 = dk.load_cfg()
                cfg2["pl_file"] = str(body.get("path", "")).strip().strip('"')
                dk.save_cfg(cfg2)
                st.pl.update(ctx=None, file="", mtime=0)
                pl_load(True)
                return {"ok": True}
            if path == "/api/ext":
                return {"ok": True, "path": str(extract_extension(dk.BASE / "uzanti"))}
            if path == "/api/ping":
                st.last_ping = time.time()
                return {"ok": True}
            if path.startswith("/api/maza/"):
                return maza_post(db, path, body, bump)
            if path == "/api/backup":
                return {"ok": True, "path": str(backup_db(db))}
            raise ValueError("naməlum yol")

    class BinaryResult(Exception):
        def __init__(self, data, ctype, name):
            self.data, self.ctype, self.name = data, ctype, name

    try:
        srv = ThreadingHTTPServer(("127.0.0.1", PORT), H)
    except OSError:
        print("Proqram artıq açıqdır - brauzer açılır.")
        if open_browser:
            webbrowser.open(f"http://127.0.0.1:{PORT}/")
        return None
    threading.Thread(target=srv.serve_forever, daemon=True).start()
    st.srv = srv
    threading.Thread(target=lambda: (do_scan(), start_lookup()), daemon=True).start()
    url = f"http://127.0.0.1:{PORT}/"
    print(f"Debitor-Kreditor ERP işləyir: {url}\nBu pəncərəni bağlasanız proqram dayanır.")
    if open_browser:
        webbrowser.open(url)
    return srv


# ----------------------------------------------------------------------------
# İnterfeys (tək HTML)
# ----------------------------------------------------------------------------
PAGE = r'''<!doctype html><html lang="az"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>M.A.Z.A ERP</title>
<style>
:root{--bg:#f4f6f9;--card:#fff;--ink:#1b2430;--mut:#5b6b7b;--line:#dfe5ea;--pri:#12395f;--acc:#1769aa;--red:#c53030;--grn:#2f855a}
@media(prefers-color-scheme:dark){:root{--bg:#10161d;--card:#18212b;--ink:#e6edf3;--mut:#9aa9b8;--line:#2a3745;--pri:#0d2b47;--acc:#4a9be0;--red:#f07070;--grn:#5fd09a}}
*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--ink);font:13px system-ui,Segoe UI,Arial}
header{background:var(--pri);color:#fff;padding:10px 16px;display:flex;gap:16px;align-items:center;flex-wrap:wrap}
header b{font-size:15px}nav{display:flex;gap:2px;flex-wrap:wrap}
nav button{background:transparent;color:#cfe0f0;border:0;padding:8px 14px;cursor:pointer;font:600 13px system-ui;border-radius:6px}
nav button.on{background:#ffffff22;color:#fff}
.bar{display:flex;gap:8px;align-items:center;flex-wrap:wrap;padding:10px 16px;background:var(--card);border-bottom:1px solid var(--line)}
.bar input[type=date],.bar input[type=text]{padding:5px 7px;border:1px solid var(--line);border-radius:5px;background:var(--bg);color:var(--ink);font:inherit}
button.b{padding:6px 12px;border:0;border-radius:6px;background:var(--acc);color:#fff;font:600 12px system-ui;cursor:pointer}
button.b.g{background:var(--grn)}button.b.s{background:var(--mut)}
main{padding:12px 16px}.card{background:var(--card);border:1px solid var(--line);border-radius:8px;overflow:auto;max-height:calc(100vh - 210px)}
table{border-collapse:collapse;width:100%}th,td{padding:5px 9px;border-bottom:1px solid var(--line);white-space:nowrap}
th{position:sticky;top:0;background:var(--pri);color:#fff;text-align:left;cursor:pointer;z-index:1}
td.n,th.n{text-align:right;font-variant-numeric:tabular-nums}tr:hover td{background:#1769aa14}
tr.risk td{color:var(--red)}tr.grp td{font-weight:700;background:var(--bg)}tr.tot td{font-weight:700;background:var(--bg);position:sticky;bottom:0}
.stat{padding:6px 16px;color:var(--mut);font-size:12px;border-top:1px solid var(--line);background:var(--card);position:fixed;bottom:0;left:0;right:0}
.tot{margin:8px 2px;font-weight:600}.pill{padding:2px 8px;border-radius:10px;font-size:11px;font-weight:600}
.p-bad{background:#c5303022;color:var(--red)}.p-warn{background:#d6992022;color:#b7791f}.p-ok{background:#2f855a22;color:var(--grn)}
.modal{position:fixed;inset:4%;background:var(--card);border:1px solid var(--line);border-radius:10px;box-shadow:0 12px 50px #0006;display:none;flex-direction:column;z-index:5}
.modal .hd{padding:10px 14px;border-bottom:1px solid var(--line);display:flex;justify-content:space-between;font-weight:700}
.modal .bd{overflow:auto;padding:0 8px 8px}label{display:block;margin:10px 0 3px;font-weight:600;color:var(--mut)}
textarea,input.f{width:100%;padding:6px;border:1px solid var(--line);border-radius:5px;background:var(--bg);color:var(--ink);font:inherit}
.grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(260px,1fr));gap:14px}.panel{padding:14px;background:var(--card);border:1px solid var(--line);border-radius:8px}
</style></head><body>
<header><b>M.A.Z.A ERP</b><nav id="nav"></nav><span id="co" style="margin-left:auto;opacity:.8"></span></header>
<div class="bar" id="bar">
 <label>Tarixdən <input type="date" id="from"></label><label>Tarixədək <input type="date" id="to"></label>
 <input type="text" id="q" placeholder="Axtar (VÖEN / ad)">
 <span id="mzsel" style="display:none"><select id="prj"><option value="">Bütün layihələr</option></select> <select id="yr"></select></span>
 <label><input type="checkbox" id="split"> Əsas / ƏDV ayrı</label>
 <button class="b" id="bShow">Göstər</button><button class="b" id="bScan">Papkaları yenilə</button>
 <button class="b s" id="bLook">VÖEN yoxla (portal)</button><button class="b g" id="bXls">Excelə ixrac</button>
</div>
<main><div class="tot" id="tot"></div><div class="card" id="view"></div></main>
<div class="stat" id="stat">Hazır</div>
<div class="modal" id="modal"><div class="hd"><span id="mt"></span><span><button class="b g" id="mx">Hərəkət Excel-i yoxdur</button> <button class="b s" id="mc">✕</button></span></div><div class="bd" id="mb"></div></div>
<script>
const T="__TOKEN__",$=id=>document.getElementById(id);
const api=async(p,o)=>{const r=await fetch(p,Object.assign({headers:{"X-UI":T,"Content-Type":"application/json"}},o));if(!r.ok&&r.status!==500)throw new Error("HTTP "+r.status);const ct=r.headers.get("content-type")||"";if(ct.includes("json")){const j=await r.json();if(j.err)throw new Error(j.err);return j}return r.blob()};
const fm=c=>c==null||c===""?"":(c/100).toLocaleString("az",{minimumFractionDigits:2,maximumFractionDigits:2}).replace(/ /g," ");
const esc=s=>String(s==null?"":s).replace(/[&<>]/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;"}[c]));
const TABS=[["deb","Debitor"],["kred","Kreditor"],["aging","Yaşlandırma"],["risk","Risk və limit"],["pl","Mənfəət-zərər"],["set","Parametrlər"],["files","Fayllar / VÖEN"],
 ["mz_pl","M-Z (baza)"],["mz_cf","Pul axını"],["mz_bs","Balans"],["mz_ratio","Ratio"],["mz_voen","VÖEN üzrə"],["mz_stock","Ehtiyat"],
 ["mz_nr","Qeyri-rezident"],["mz_map","Mapping"],["mz_cap","Kapital / ƏV"],["mz_imp","İdxal"]];
let tab="deb",data=null,sortI=-1,sortD=1,rev=-1;
$("nav").innerHTML=TABS.map(t=>`<button data-t="${t[0]}">${t[1]}</button>`).join("");
document.querySelectorAll("#nav button").forEach(b=>b.onclick=()=>{tab=b.dataset.t;mzData=null;load()});
const qs=()=>{const p=new URLSearchParams();if($("from").value)p.set("from",$("from").value);if($("to").value)p.set("to",$("to").value);return p};
// DK cədvəli: [başlıq, əsas indeksi, ƏDV indeksi] (cəmi = əsas+ƏDV)
const G=[["Əvvəlki qalıq",4,5],["Qaimə",6,7],["Qaytarılma",8,9],["Ödəniş",10,11],["Qaytarılan pul",12,13],["Qalıq",14,15]];
function table(cols,rows,opt={}){ // cols: [title,fn(row)->value,type:'n'|'t', raw]
 let h="<table><tr>"+cols.map((c,i)=>`<th class="${c[2]==="n"?"n":""}" data-i="${i}">${c[0]}${sortI===i?(sortD>0?" ▲":" ▼"):""}</th>`).join("")+"</tr>";
 rows.forEach((r,ri)=>{h+=`<tr data-r="${ri}" class="${opt.cls?opt.cls(r):""}">`+cols.map(c=>{const v=c[1](r);return `<td class="${c[2]==="n"?"n":""}">${c[2]==="n"?fm(v):(c[3]?v:esc(v))}</td>`}).join("")+"</tr>"});
 if(opt.total){h+='<tr class="tot">'+cols.map((c,i)=>`<td class="${c[2]==="n"?"n":""}">${i===0?"YEKUN":c[2]==="n"&&c[4]!==0?fm(rows.reduce((a,r)=>a+(+c[1](r)||0),0)):""}</td>`).join("")+"</tr>"}
 return h+"</table>"}
function filt(rows,nameFn){const q=$("q").value.toLowerCase().trim();return q?rows.filter(r=>nameFn(r).toLowerCase().includes(q)):rows}
function bind(rows,cols,onRow){document.querySelectorAll("#view th").forEach(th=>th.onclick=()=>{const i=+th.dataset.i;sortD=sortI===i?-sortD:1;sortI=i;render()});
 if(onRow)document.querySelectorAll("#view tr[data-r]").forEach(tr=>tr.onclick=()=>onRow(rows[+tr.dataset.r]))}
function sorted(rows,cols){if(sortI<0||!cols[sortI])return rows;const f=cols[sortI][1];return [...rows].sort((a,b)=>{const x=f(a),y=f(b);return(typeof x==="number"&&typeof y==="number"?x-y:String(x).localeCompare(String(y),"az"))*sortD})}
function render(){
 $("mzsel").style.display=tab.startsWith("mz_")?"inline":"none";
 if(tab.startsWith("mz_")){mzRender();return}
 const v=$("view");$("bar").style.display=["set","files","pl"].includes(tab)?"none":"flex";
 document.querySelectorAll("#nav button").forEach(b=>b.classList.toggle("on",b.dataset.t===tab));
 if(tab==="deb"||tab==="kred"){if(!data||!data.rows)return;const sp=$("split").checked;
  const cols=[["VÖEN",r=>r[0],"t"],["Tərəfin adı",r=>r[1],"t"],["ƏDV ödəyicisi",r=>r[2],"t"],["Riskli",r=>r[3],"t"]];
  G.forEach(g=>{if(sp){cols.push([g[0]+" əsas",r=>r[g[1]],"n"],[g[0]+" ƏDV",r=>r[g[2]],"n"])}else cols.push([g[0],r=>r[g[1]]+r[g[2]],"n"])});
  if(sp)cols.push(["Qalıq cəmi",r=>r[16],"n"]);
  const rows=sorted(filt(data.rows,r=>r[0]+" "+r[1]),cols);v.innerHTML=table(cols,rows,{total:1,cls:r=>r[3]==="Bəli"?"risk":""});bind(rows,cols,r=>openStmt(r[0],r[1]));
  const s=filt(data.rows,r=>r[0]+" "+r[1]);$("tot").textContent=(tab==="deb"?"Debitor":"Kreditor")+" yekun: "+fm(s.reduce((a,r)=>a+r[16],0))+"  ("+s.length+" tərəf)"; }
 else if(tab==="aging"){if(!data||!data.deb)return;const L=data.labels;let out="";
  [["deb","Debitor yaşlandırma (açıq borc, "+data.as_of+" tarixinə)"],["kred","Kreditor yaşlandırma"]].forEach(([k,t])=>{
   const cols=[["VÖEN",r=>r.voen,"t"],["Tərəfin adı",r=>r.name,"t"],...L.map((l,i)=>[l,r=>r.buckets[i],"n"]),["Cəmi borc",r=>r.total,"n"]];
   if(k==="deb"&&data.allow.some(x=>+x))cols.push(["Ehtiyat",r=>r.reserve,"n"]);
   out+=`<h3 style="margin:10px 8px">${t}</h3>`+table(cols,filt(data[k],r=>r.voen+" "+r.name),{total:1})});
  v.innerHTML=out;$("tot").textContent="Qalıq ən son borc sənədlərinə (sondan əvvələ) bölünür."}
 else if(tab==="risk"){if(!data||!data.rows)return;
  const cols=[["VÖEN",r=>r.voen,"t"],["Tərəfin adı",r=>r.name,"t"],["Debitor borc",r=>r.debt,"n"],["Limit",r=>r.limit,"n",0,0],["Boş limit",r=>r.free,"n",0,0],["İstifadə %",r=>r.use==null?"":r.use.toFixed(1)+"%","t"],["Gecikmiş borc",r=>r.over,"n"],
   ["Status",r=>`<span class="pill ${r.status==="Normal"?"p-ok":r.status.includes("aşılıb")?"p-bad":"p-warn"}">${r.status}</span>`,"t",1]];
  const rows=sorted(filt(data.rows,r=>r.voen+" "+r.name),cols);v.innerHTML=table(cols,rows,{total:1});bind(rows,cols,r=>openStmt(r.voen,r.name));
  $("tot").textContent=`Risk altında: ${data.rows.filter(r=>r.status!=="Normal").length} / ${data.rows.length} debitor. Gecikmiş = ${data.settings.credit_overdue_days} gündən köhnə borc.`}
 else if(tab==="set")settingsView();else if(tab==="files")filesView();else if(tab==="pl")plView();
}
async function load(){ if(tab.startsWith("mz_")){mzLoad();return}
 if(tab==="set"||tab==="files"||tab==="pl"){render();return}
 const p=qs();try{
  if(tab==="deb"||tab==="kred"){p.set("side",tab);data=await api("/api/balances?"+p)}
  else if(tab==="aging")data=await api("/api/aging?"+p);else if(tab==="risk")data=await api("/api/risk?"+p);
  render()}catch(e){$("view").innerHTML="<p style='padding:14px;color:var(--red)'>"+esc(e.message)+"</p>"}}
async function openStmt(voen,name){const d=await api("/api/statement?voen="+encodeURIComponent(voen)+"&"+qs());
 const H=["Tarix","Sənəd №","Əməliyyat","E-qaimə əsas","E-qaimə ƏDV","Qaytarılma əsas","Qaytarılma ƏDV","Ödəniş əsas","Ödəniş ƏDV","Qaytarılan pul əsas","Qaytarılan pul ƏDV","Qalıq əsas","Qalıq ƏDV","Qeyd"];
 const cols=H.map((h,i)=>[h,r=>i===0?r[0].split("-").reverse().join("."):i<3||i===13?(r[i]||""):r[i]?r[i]:"",i<3||i===13?"t":"n"]);
 $("mt").textContent=(d.name||name)+" ("+voen+")";$("mb").innerHTML=table(cols,d.rows,{});$("modal").style.display="flex";$("mx").style.display="none"}
$("mc").onclick=()=>$("modal").style.display="none";
// ================= M.A.Z.A =================
const MZ_REP={mz_pl:"pl",mz_cf:"cf",mz_bs:"bs",mz_ratio:"ratio",mz_voen:"voen",mz_stock:"stock"};
const MZ_EDIT=["mz_nr","mz_map","mz_cap","mz_imp"];
const MZ_SUBS={
 mz_nr:[["nonres","Qeyri-rezident idxal / xidmət"]],
 mz_map:[["map_class","Qaimə təsnifatı"],["map_bank","Bank qaydaları"],["map_item","Mal adları"],["unmapped","Təsnifatsızlar"]],
 mz_cap:[["equity","Kapital / təsisçi borcu"],["opening","Açılış qalıqları"],["assets","Əsas vəsaitlər"],["projects","Layihələr"],["settings","Vergi ayarları"]]};
const KIND_CF={"Kredit əsas məbləğ":"Maliyyə","Kredit faiz":"Maliyyə","Təsisçi":"Maliyyə","Hesablararası":"Daxili"};
let mzInit=null,mzData=null,mzRep=null,mzSub={mz_nr:"nonres",mz_map:"map_class",mz_cap:"equity"};
const f2=v=>v===""||v==null?"":(typeof v==="number"?v.toLocaleString("az",{minimumFractionDigits:2,maximumFractionDigits:2}):esc(v));
const mzQs=()=>{const p=qs();if($("prj").value)p.set("project",$("prj").value);if($("yr").value)p.set("year",$("yr").value);return p};
const mzErr=e=>"<p style='padding:14px;color:var(--red)'>"+esc(e.message)+"</p>";
async function mzBoot(){try{mzInit=await api("/api/maza/init");const pv=$("prj").value,yv=$("yr").value||String(new Date().getFullYear());
 $("prj").innerHTML='<option value="">Bütün layihələr</option>'+mzInit.projects.map(p=>`<option ${p===pv?"selected":""}>${esc(p)}</option>`).join("");
 $("yr").innerHTML=mzInit.years.map(y=>`<option ${String(y)===yv?"selected":""}>${y}</option>`).join("")}catch(e){}}
async function mzLoad(){if(MZ_REP[tab]){try{mzData=await api("/api/maza/"+MZ_REP[tab]+"?"+mzQs());mzRender()}catch(e){$("view").innerHTML=mzErr(e)}}else mzRender()}
function mzTable(d){let h="<table><tr>"+d.cols.map(c=>`<th class="${c[1]?"n":""}">${esc(c[0])}</th>`).join("")+"</tr>";
 d.rows.forEach(r=>{h+=`<tr class="${r.k==="grp"?"grp":r.k==="tot"?"tot":r.k==="risk"?"risk":""}">`+r.c.map((x,i)=>`<td class="${d.cols[i][1]?"n":""}">${d.cols[i][1]?f2(x):esc(x)}</td>`).join("")+"</tr>"});return h+"</table>"}
function mzRender(){
 $("mzsel").style.display="inline";
 document.querySelectorAll("#nav button").forEach(b=>b.classList.toggle("on",b.dataset.t===tab));
 $("bar").style.display=MZ_EDIT.includes(tab)?"none":"flex";
 if(MZ_EDIT.includes(tab)){$("tot").textContent="";mzEditView();return}
 if(!mzData)return;
 const q=$("q").value.toLowerCase().trim();
 const d=q&&["mz_voen","mz_stock"].includes(tab)?Object.assign({},mzData,{rows:mzData.rows.filter(r=>r.k==="tot"||(r.c[0]+" "+r.c[1]).toLowerCase().includes(q))}):mzData;
 $("view").innerHTML=mzTable(d);$("tot").textContent=mzData.note||""}
async function mzEditView(){
 if(tab==="mz_imp"){mzImpView();return}
 const subs=MZ_SUBS[tab],cur=mzSub[tab],v=$("view");
 v.innerHTML=`<div style="padding:10px 12px;display:flex;gap:6px;flex-wrap:wrap">`+(subs.length>1?subs.map(s=>`<button class="b ${s[0]===cur?"":"s"}" data-s="${s[0]}">${s[1]}</button>`).join(""):"")+`</div><div id="mzh"></div>`;
 v.querySelectorAll("button[data-s]").forEach(b=>b.onclick=()=>{mzSub[tab]=b.dataset.s;mzEditView()});
 const host=$("mzh");host.innerHTML="<p style='padding:12px;color:var(--mut)'>Yüklənir…</p>";
 try{if(cur==="unmapped")await mzUnmapped(host);else if(cur==="settings")await mzSettings(host);else await mzCrud(cur,host)}catch(e){host.innerHTML=mzErr(e)}}
async function mzCrud(tb,host){const d=await api("/api/maza/list?table="+tb),F=d.fields;
 let h=F.filter(f=>f.o).map(f=>`<datalist id="dl_${f.k}">${f.o.map(o=>`<option value="${esc(o)}">`).join("")}</datalist>`).join("");
 h+=`<form id="mzf" style="display:grid;grid-template-columns:repeat(auto-fit,minmax(170px,1fr));gap:8px;padding:10px 12px"><input type="hidden" name="id">`+
  F.map(f=>`<label style="margin:0">${esc(f.l)}<input class="f" name="${f.k}" type="${f.t==="date"?"date":"text"}" ${f.o?`list="dl_${f.k}"`:""}></label>`).join("")+
  `<div style="align-self:end"><button type="submit" class="b">Saxla</button> <button type="button" class="b s" id="mzn">Təmizlə</button></div></form>`;
 h+="<table><tr>"+F.map(f=>`<th class="${f.t==="num"?"n":""}">${esc(f.l)}</th>`).join("")+"<th></th></tr>"+
  d.rows.map((r,i)=>"<tr>"+F.map(f=>`<td class="${f.t==="num"?"n":""}">${f.t==="num"?(f.k==="prio"?esc(r[f.k]):f2(+r[f.k])):esc(r[f.k])}</td>`).join("")+`<td><button class="b s" data-e="${i}">Düzəliş</button> <button class="b s" data-d="${r.id}">Sil</button></td></tr>`).join("")+"</table>";
 host.innerHTML=h;const form=host.querySelector("#mzf");
 form.onsubmit=async e=>{e.preventDefault();const row={};new FormData(form).forEach((v,k)=>row[k]=v);if(row.id)row.id=+row.id;else delete row.id;
  try{await api("/api/maza/save",{method:"POST",body:JSON.stringify({table:tb,row})});if(tb==="projects")mzBoot();mzEditView()}catch(er){alert(er.message)}};
 host.querySelector("#mzn").onclick=()=>form.reset();
 host.querySelectorAll("button[data-e]").forEach(b=>b.onclick=()=>{const r=d.rows[+b.dataset.e];form.elements.id.value=r.id;F.forEach(f=>{form.elements[f.k].value=r[f.k]==null?"":r[f.k]});form.scrollIntoView({block:"nearest"})});
 host.querySelectorAll("button[data-d]").forEach(b=>b.onclick=async()=>{if(!confirm("Silinsin?"))return;await api("/api/maza/delete",{method:"POST",body:JSON.stringify({table:tb,id:+b.dataset.d})});if(tb==="projects")mzBoot();mzEditView()})}
async function mzUnmapped(host){
 const [u,mc,mb,it]=await Promise.all([api("/api/maza/unmapped"),api("/api/maza/list?table=map_class"),api("/api/maza/list?table=map_bank"),api("/api/maza/items")]);
 const fo=(F,k)=>(F.find(f=>f.k===k)||{}).o||[];
 const sel=(id,arr,def)=>`<select id="${id}">${arr.map(o=>`<option ${o===def?"selected":""}>${esc(o)}</option>`).join("")}</select>`;
 const CLS=fo(mc.fields,"cls"),PL=["",...fo(mc.fields,"pl_line")],KIND=fo(mb.fields,"kind");
 let h=`<h3 style="margin:10px 12px">Təsnifatsız qaimələr (VÖEN üzrə)</h3>`;
 h+=u.inv.length?`<table><tr><th>Tərəf</th><th>VÖEN</th><th>Ad</th><th class="n">Sətir</th><th class="n">Cəmi</th><th>Nümunə</th><th>Sinif</th><th>P&amp;L maddəsi</th><th></th></tr>`+
  u.inv.map((r,i)=>`<tr><td>${r.dir==="G"?"Gələn":"Gedən"}</td><td>${esc(r.voen)}</td><td>${esc(r.name)}</td><td class="n">${r.n}</td><td class="n">${f2(r.total)}</td><td>${esc(r.sample)}</td><td>${sel("ic"+i,CLS,r.dir==="S"?"gelir":"xerc")}</td><td>${sel("ip"+i,PL,"")}</td><td><button class="b" data-k="inv" data-i="${i}">Qayda yarat</button></td></tr>`).join("")+"</table>":"<p style='padding:0 12px;color:var(--grn)'>✔ Hamısı təsnif olunub</p>";
 h+=`<h3 style="margin:14px 12px 6px">Təsnifatsız bank hərəkətləri</h3>`;
 h+=u.bank.length?`<table><tr><th>VÖEN</th><th class="n">Sətir</th><th class="n">Mədaxil</th><th class="n">Məxaric</th><th>Təyinat (nümunə)</th><th>Hərəkət növü</th><th></th></tr>`+
  u.bank.map((r,i)=>`<tr><td>${esc(r.voen)}</td><td class="n">${r.n}</td><td class="n">${f2(r.in)}</td><td class="n">${f2(r.out)}</td><td>${esc(r.sample)}</td><td>${sel("bk"+i,KIND,r.in>r.out?"Debitor ödəmə":"Kreditor ödəmə")}</td><td><button class="b" data-k="bank" data-i="${i}">Qayda yarat</button></td></tr>`).join("")+"</table>":"<p style='padding:0 12px;color:var(--grn)'>✔ Hamısı təsnif olunub</p>";
 h+=`<h3 style="margin:14px 12px 6px">Mal adlarının uyğunlaşdırılması (ehtiyat dövriyyəsi üçün məcburi)</h3><p style="padding:0 12px;color:var(--mut)">Eyni malın alış və satış qaimələrindəki fərqli adlarını bir vahid adda birləşdirin. Dəyişməsə vahid adı eyni saxlayıb "Yarat" basın.</p>`;
 h+=it.items.length?`<table><tr><th>Qaimədəki ad</th><th>Hərəkət</th><th class="n">Alış sətri</th><th class="n">Satış sətri</th><th class="n">Əsas məbləğ</th><th>Vahid ad</th><th>Vahid</th><th></th></tr>`+
  it.items.map((r,i)=>`<tr><td>${esc(r.item)}</td><td>${r.role}</td><td class="n">${r.g}</td><td class="n">${r.s}</td><td class="n">${f2(r.total)}</td><td><input class="f" id="is${i}" value="${esc(r.item)}" style="min-width:240px"></td><td><input class="f" id="iu${i}" value="${esc(r.unit)}" style="width:70px"></td><td><button class="b" data-k="item" data-i="${i}">Yarat</button></td></tr>`).join("")+"</table>"+(it.more?`<p style="padding:6px 12px;color:var(--mut)">… və daha ${it.more} ad</p>`:""):"<p style='padding:0 12px;color:var(--grn)'>✔ Hamısı uyğunlaşdırılıb</p>";
 host.innerHTML=h;
 host.querySelectorAll("button[data-k]").forEach(b=>b.onclick=async()=>{const i=+b.dataset.i,k=b.dataset.k;let tb,row;
  if(k==="inv"){const r=u.inv[i];tb="map_class";row={mtype:"voen",pattern:r.voen,side:r.dir,cls:$("ic"+i).value,pl_line:$("ip"+i).value,project:"",prio:3}}
  else if(k==="bank"){const r=u.bank[i];tb="map_bank";const kind=$("bk"+i).value;let mt="voen",pat=r.voen;
   if(!pat){mt="text";pat=prompt("Təyinatda axtarılan açar söz:",(r.sample||"").slice(0,25));if(!pat)return}
   row={mtype:mt,pattern:pat,kind,cf_section:KIND_CF[kind]||"Əməliyyat",pl_line:"",project:"",prio:3}}
  else{const r=it.items[i];tb="map_item";row={raw:r.item,std:$("is"+i).value||r.item,unit:$("iu"+i).value}}
  b.disabled=true;try{await api("/api/maza/save",{method:"POST",body:JSON.stringify({table:tb,row})});mzUnmapped(host)}catch(er){alert(er.message);b.disabled=false}})}
async function mzSettings(host){const d=await api("/api/maza/init"),s=d.settings;
 const rows=[["rate","Mənfəət vergisi dərəcəsi (%)",+(s.rate*100).toFixed(4)],["loss1","Yığılmış zərər 2023 (AZN)",s.loss1],["loss2","Yığılmış zərər 2024 (AZN)",s.loss2],
  ["cap1","Bina tikili qurğu təmir xərci üzrə limit (AZN, 0 = limitsiz)",s.cap1],["cap2","Nəqliyyat ehtiyat hissələri üzrə limit (AZN, 0 = limitsiz)",s.cap2],["own_voen","Şirkətin VÖEN-i (portal ixracında gələn/gedən təyini)",s.own_voen||""]];
 host.innerHTML=`<div style="padding:14px;max-width:560px">`+rows.map(r=>`<label>${r[1]}</label><input class="f" id="zs_${r[0]}" value="${esc(r[2])}">`).join("")+`<p><button class="b" id="zsv">Yadda saxla</button> <span id="zsm"></span></p></div>`;
 $("zsv").onclick=async()=>{const body={};rows.forEach(r=>body[r[0]]=$("zs_"+r[0]).value);try{await api("/api/maza/settings",{method:"POST",body:JSON.stringify(body)});$("zsm").textContent="✔ Yadda saxlanıldı"}catch(e){$("zsm").textContent=e.message}}}
async function mzImpView(){const v=$("view");
 try{await mzBoot();const [c,ck]=await Promise.all([api("/api/maza/init"),api("/api/maza/checks")]);
  const CN={inv:"Qaimə sətirləri",bank:"Bank hərəkətləri",nonres:"Qeyri-rezident sənədlər",payroll:"Əmək haqqı sətirləri",assets:"Əsas vəsaitlər",map_class:"Qaimə qaydaları",map_bank:"Bank qaydaları",projects:"Layihələr",equity:"Kapital hərəkətləri"};
  v.innerHTML=`<div style="padding:14px"><div class="grid"><div class="panel"><b>Excel / xlsb idxalı</b><label>Fayl və ya qovluq</label><input class="f" id="zp" placeholder="${esc(c.folder)}">
  <p style="color:var(--mut);font-size:12px">Boş qoysanız standart qovluq oxunur. Hər vərəq başlığına görə tanınır (portal ixracı, gələn/gedən qaimə, bank çıxarışı, ecustom, qeyri-rezident xidmət, əmək haqqı).</p>
  <p><button class="b" id="zi">Oxu</button> <button class="b s" id="zc">Yenidən təsnifləşdir</button> <button class="b s" id="zt">Excel şablonlarını yarat</button></p>
  <p><button class="b s" id="zs">Debitor / Kreditor cədvəlinə ötür</button></p><p id="zm" style="color:var(--mut)"></p></div>
  <div class="panel"><b>Bazada</b>${Object.keys(CN).map(k=>`<p style="margin:4px 0">${CN[k]}: <b>${c.counts[k]}</b></p>`).join("")}</div>
  <div class="panel"><b>Yoxlama</b>${ck.rows.map(r=>`<p style="margin:4px 0">${r.ok?"✔":"⚠"} ${esc(r.label)}: <b>${r.n}</b> <span style="color:var(--mut);font-size:11px">${r.ok?"":esc(r.hint)}</span></p>`).join("")}</div></div>
  <div id="zr" style="margin-top:12px">${mzRep===null?"":mzRep.length?`<div class="card"><table><tr><th>Mənbə (fayl | vərəq)</th><th>Növ</th><th class="n">Yeni sətir</th><th class="n">Dublikat (keçildi)</th></tr>${mzRep.map(r=>`<tr><td>${esc(r[0])}</td><td>${esc(r[1])}</td><td class="n">${r[2]}</td><td class="n">${r[3]}</td></tr>`).join("")}</table></div>`:"<p style='color:var(--red)'>Tanınan vərəq tapılmadı.</p>"}</div></div>`;
  $("zi").onclick=async()=>{$("zm").textContent="Oxunur… (böyük fayllar bir neçə dəqiqə çəkə bilər)";
   try{const r=await api("/api/maza/import",{method:"POST",body:JSON.stringify({path:$("zp").value.trim()})});mzRep=r.report;mzImpView()}catch(e){$("zm").textContent="Xəta: "+e.message}};
  $("zc").onclick=async()=>{await api("/api/maza/classify",{method:"POST",body:"{}"});mzImpView()};
  $("zt").onclick=async()=>{const r=await api("/api/maza/templates",{method:"POST",body:"{}"});$("zm").textContent="Yaradıldı: "+r.files.join(", ")};
  $("zs").onclick=async()=>{if(!confirm("Qaimə və bank hərəkətləri Debitor/Kreditor cədvəlinə köçürüləcək. Eyni məlumat uzantıdan və ya papkalardan da gəlirsə, qalıqlar iki dəfə sayılar. Davam edilsin?"))return;
   const r=await api("/api/maza/sync",{method:"POST",body:"{}"});$("zm").textContent="✔ "+r.rows+" sətir ötürüldü"}
 }catch(e){v.innerHTML=mzErr(e)}}
// ================= /M.A.Z.A =================
let plYear="";
async function plView(){const d=await api("/api/pl"+(plYear?"?year="+plYear:""));
 let h=`<div style="padding:14px"><div class="grid"><div class="panel"><b>Excel iş kitabı (mənfəət-zərər)</b><label>Fayl yolu</label><input class="f" id="pf" value="${esc(d.file)}">
 <p><button class="b" id="pfs">Saxla və oxu</button> <span style="color:var(--mut)">${esc(d.msg)}</span></p></div></div>`;
 if(d.ready){const cols=[["Maddə",r=>r.label,"t"],...d.months.map((m,i)=>[m,r=>r.m[i],"n"]),["İllik",r=>r.cum,"n"]];
  h+=`<p><label style="display:inline">İl <select id="py">${d.years.map(y=>`<option ${y===d.year?"selected":""}>${y}</option>`).join("")}</select></label> <button class="b g" id="pxl">Excelə ixrac</button></p>`;
  if(d.rec)h+=`<p style="color:var(--mut)">Excel-in "pl" vərəqi (${d.rec.year}) ilə yoxlama: ${d.rec.checked} xanadan ${d.rec.checked-d.rec.n} uyğundur, ${d.rec.n} fərq.`+(d.rec.n?`<details><summary>Fərqlər</summary><pre style="font-size:11px">${esc(d.rec.first.map(x=>x.join(" | ")).join("\n"))}</pre></details>`:"")+`</p>`;
  h+=`<div class="card" style="max-height:calc(100vh - 380px)">${table(cols,d.rows,{cls:r=>r.kind==="item"?"":"grp"})}</div>`}
 $("view").innerHTML=h+"</div>";$("tot").textContent="";
 $("pfs").onclick=async()=>{await api("/api/pl/file",{method:"POST",body:JSON.stringify({path:$("pf").value})});plView()};
 if(d.ready){$("py").onchange=e=>{plYear=e.target.value;plView()};
  $("pxl").onclick=async()=>{const b=await api("/api/pl/export?year="+d.year);const a=document.createElement("a");a.href=URL.createObjectURL(b);a.download="menfeet_zerer_"+d.year+".xlsx";a.click()}}
 if(d.busy)setTimeout(()=>{if(tab==="pl")plView()},1500)}
async function settingsView(){const s=await api("/api/settings"),R=["0–30","31–60","61–90","91–180","180+"];
 $("view").innerHTML=`<div style="padding:14px"><div class="grid"><div class="panel"><b>Kredit limiti</b><label>Ümumi limit (AZN, 0 = yoxdur)</label><input class="f" id="s1" value="${s.credit_default}">
 <label>Gecikmə həddi (gün; bundan köhnə borc "gecikmiş")</label><input class="f" id="s2" value="${s.credit_overdue_days}">
 <label>Fərdi limitlər (hər sətir: VÖEN;limit)</label><textarea id="s3" rows="8">${esc(s.credit_limits)}</textarea></div>
 <div class="panel"><b>Debitor ehtiyatı (şübhəli borc), %</b>${R.map((r,i)=>`<label>${r} gün</label><input class="f" id="a${i}" value="${s.allow[i]}">`).join("")}</div></div>
 <p><button class="b" id="sv">Yadda saxla</button> <span id="sm"></span></p></div>`;
 $("sv").onclick=async()=>{await api("/api/settings",{method:"POST",body:JSON.stringify({credit_default:+$("s1").value||0,credit_overdue_days:+$("s2").value||60,credit_limits:$("s3").value,allow:[0,1,2,3,4].map(i=>+$("a"+i).value||0)})});$("sm").textContent="✔ Yadda saxlanıldı"}}
async function filesView(){const s=await api("/api/status");$("view").innerHTML=`<div style="padding:14px"><div class="grid"><div class="panel"><b>Papkalar</b>
 <label>Qaimələr papkası</label><input class="f" id="f1" value="${esc(s.qaime)}"><label>Ödənişlər papkası</label><input class="f" id="f2" value="${esc(s.odenis)}">
 <p><button class="b" id="fs">Saxla və yenilə</button> <button class="b s" id="ft">Excel şablonlarını yarat</button> <button class="b s" id="fe">Chrome uzantısını çıxar</button> <button class="b s" id="fb">Bazanın ehtiyat nüsxəsi</button></p><p id="fm" style="color:var(--mut)"></p></div>
 <div class="panel"><b>Vəziyyət</b><p>Şirkət (VÖEN): ${esc(s.company||"hələ təyin olunmayıb")}</p><p>Bazada ${s.rows} əməliyyat</p><p>${esc(s.scan)}</p><p>${esc(s.incoming)}</p><p style="color:var(--mut);font-size:11px">${esc(s.db)}</p>
 ${s.errors.length?"<b style='color:var(--red)'>Oxunmayan sətirlər:</b><pre style='font-size:11px;max-height:160px;overflow:auto'>"+esc(s.errors.map(e=>e.join(" | ")).join("\n"))+"</pre>":""}</div></div></div>`;
 $("fs").onclick=async()=>{await api("/api/folders",{method:"POST",body:JSON.stringify({qaime:$("f1").value,odenis:$("f2").value})});$("fm").textContent="✔ Yeniləndi"};
 $("fb").onclick=async()=>{const r=await api("/api/backup",{method:"POST",body:"{}"});$("fm").textContent="Ehtiyat nüsxə: "+r.path};
 $("fe").onclick=async()=>{const r=await api("/api/ext",{method:"POST",body:"{}"});$("fm").textContent="Uzantı qovluğu: "+r.path+" (chrome://extensions → Load unpacked)"};
 $("ft").onclick=async()=>{const r=await api("/api/templates",{method:"POST",body:"{}"});$("fm").textContent="Yaradıldı: "+r.files.join(", ")}}
$("prj").onchange=load;$("yr").onchange=load;
$("bShow").onclick=load;$("q").oninput=render;$("split").onchange=render;
$("bScan").onclick=async()=>{$("stat").textContent="Papkalar oxunur…";await api("/api/scan",{method:"POST",body:"{}"});load()};
$("bLook").onclick=()=>api("/api/lookup",{method:"POST",body:JSON.stringify({force:false})});
$("bXls").onclick=async()=>{if(tab.startsWith("mz_")){const b=await api("/api/maza/export?"+mzQs());const a=document.createElement("a");a.href=URL.createObjectURL(b);a.download="maza_hesabat_"+($("yr").value||"")+".xlsx";a.click();return}
 const b=await api("/api/export?"+qs());const a=document.createElement("a");a.href=URL.createObjectURL(b);a.download="debitor_kreditor.xlsx";a.click()};
const t0=new Date(),iso=d=>d.toISOString().slice(0,10);$("to").value=iso(t0);
async function poll(){try{const s=await api("/api/status");$("co").textContent=s.company?"VÖEN: "+s.company:"";
 let t=`Bazada ${s.rows} əməliyyat. `+(s.incoming||"");const l=s.lookup;if(l)t+=` | VÖEN portaldan: ${l.done}/${l.total}${l.err?" (xəta "+l.err+": "+l.msg+")":""}${l.fin?" ✔":""}`;$("stat").textContent=t;
 if(s.rev!==rev){const first=rev<0;rev=s.rev;if(tab==="pl")render();else if(!["set","files"].includes(tab)&&!MZ_EDIT.includes(tab))load()}}catch(e){$("stat").textContent="Proqrama qoşulmaq olmur (bağlanıb?)"}setTimeout(poll,2000)}
load();poll();mzBoot();
</script></body></html>'''


# ----------------------------------------------------------------------------
# ERP modulu: Mənfəət-zərər hesabatı (Excel iş kitabından; erp_program/core.js-in Python tərcüməsi)
# ----------------------------------------------------------------------------
PL_MONTHS = ['Yanvar', 'Fevral', 'Mart', 'Aprel', 'May', 'İyun', 'İyul', 'Avqust', 'Sentyabr', 'Oktyabr', 'Noyabr', 'Dekabr']
PL_MONTH_COLS = ['D', 'G', 'J', 'M', 'P', 'S', 'V', 'Y', 'AB', 'AE', 'AH', 'AK']
PL_ASSET_SHEET = 'esas vesait kateqoriyalar uzre'
PL_PAY = {'salary': 'Əmək haqqı xərci', 'dsmf': 'DSMF — işəgötürən',
          'unemployment': 'İşsizlik — işəgötürən', 'its': 'İTS — işəgötürən'}

# amt/lab/lab2/y/m = sütun hərfləri; from = ilk sətir; head = başlıq yoxlaması; spread = ilə bölünür
PL_SOURCES = {
    'lim': dict(sheet='Liman icarə gəlirləri və aşırma', head={'E2': 'Məbləğ (AZN)'}, amt='E', lab='F', y='H', m='G', frm=3),
    'ged': dict(sheet='geden qaimə', head={'P1': 'Malın ƏDV-siz ümumi dəyəri'}, amt='P', lab='F', y='H', m='G', frm=2),
    'gel': dict(sheet='Gələn qaimə', head={'Z1': 'Əsas məbləğ'}, amt='Z', lab='G', y='I', m='H', frm=2),
    'gemi': dict(sheet='Gəmi icarə xərcləri', head={'Y5': 'Yekun'}, amt='Y', lab='AA', y='AC', m='AB', frm=6),
    'ecu': dict(sheet='ecustom', head={'V1': 'Əsas məbləğ'}, amt='V', lab='T', y='W', m='X', frm=2),
    'bank': dict(sheet='bank cixaris', head={'H1': 'Məxaric'}, amt='H', lab='L', lab2='K', y='Q', m='R', frm=2, inn='G'),
    'ev': dict(sheet='əsas vəsait cədvəli', head={'T2': 'Məbləğ'}, amt='T', lab='W', frm=3, spread=True, need_label=True),
    'sv': dict(sheet='Şəxsi vərəqə', head={'E1': 'Miqdar (Manat)'}, amt='E', lab='K', y='J', m='I', frm=2, spread=True),
}
PL_REV = ['lim', 'ged']
PL_COGS = ['gemi', 'gel']
PL_OPEX = ['gemi', 'gel', 'ecu', 'bank', 'ev', 'sv']


def _rng(a, b):
    return list(range(a, b + 1))


def _item(r, label, src, **kw):
    d = dict(r=r, label=label, src=src, kind='item')
    d.update(kw)
    return d


PL_LINES = [
    dict(r=6, label='SATIŞ', kind='group', of=_rng(7, 12)),
    _item(7, 'Aşırma xidməti', PL_REV), _item(8, 'Metal gəlirləri', PL_REV), _item(9, 'Gəmi icarə gəlirləri', PL_REV),
    _item(10, 'Liman Gəlirləri', PL_REV), _item(11, 'Anbar İcarə', PL_REV), _item(12, 'Digər', PL_REV),
    dict(r=14, label='MAYA DƏYƏRİ', kind='group', of=_rng(15, 17)),
    _item(15, 'Metal çıxarılması və daşınması ilə bağlı xərclər', PL_COGS),
    _item(16, 'Gəmi gəlirləriləri ilə bağlı xərclər', PL_COGS),
    _item(17, 'İşçi əmək haqqı xərcləri', PL_COGS),
    dict(r=20, label='BRUT QAZANC', kind='total'),
    dict(r=22, label='Toplam Əməliyyat xərcləri', kind='group', of=_rng(23, 44) + [24.5]),
    _item(23, PL_PAY['salary'], ['eh']), _item(24, PL_PAY['dsmf'], ['eh']),
    _item(24.5, PL_PAY['unemployment'], ['eh']), _item(25, PL_PAY['its'], ['eh']),
    _item(26, 'Daşınma xərci', PL_OPEX),
    _item(27, 'Bina tikili qurgu təmir xərci', PL_OPEX, cap=0),
    _item(28, 'Nəqliyyat ehtiyyat hissələri', PL_OPEX, cap=1),
    _item(29, 'Avadanlıq təmir xərcləri', PL_OPEX), _item(30, 'Mühafizə xərci', PL_OPEX),
    _item(31, 'Bilavasitə xərclər', PL_OPEX), _item(32, 'Enerji xərci', PL_OPEX),
    _item(33, 'Reklam xərci', PL_OPEX), _item(34, 'Ərzaq xərci', PL_OPEX), _item(35, 'Sığorta xərci', PL_OPEX),
    _item(36, 'Digər xərc', PL_OPEX), _item(37, 'Rabitə xərci', PL_OPEX), _item(38, 'İcarə xərci', PL_OPEX, extra=True),
    _item(39, 'Su və kanalizasiya xərcləri', PL_OPEX), _item(40, 'Qaz xərci', PL_OPEX),
    _item(41, 'Yanacaq xərci', PL_OPEX), _item(42, 'Əmlak vergisi', PL_OPEX),
    _item(43, 'Torpaq icarə xərcləri', PL_OPEX), _item(44, 'Torpaq vergisi', PL_OPEX),
    dict(r=46, label='AMORTIZASIYA XƏRCLƏRİ', kind='group', of=_rng(47, 52)),
    _item(47, 'Bina tikili və qurğular Amortizasiya', ['ev']), _item(48, 'Maşın və avadanlıqılar Amortizasiya', ['ev']),
    _item(49, 'Nəqliyyat vasitələri Amortizasiya', ['ev']), _item(50, 'Qeyri maddi aktivlər Amortizasiya', ['ev']),
    _item(51, 'Digər əsas vəsaitlər Amortizasiya', ['ev']),
    _item(52, 'Yüksək texnologiya məhsulları Amortizasiya', ['ev']),
    dict(r=55, label='MALİYYƏ XƏRCLƏRİ', kind='group', of=[56]),
    _item(56, 'Bank xidmət haqqı', ['bankK']),
    dict(r=58, label='Vergidən öncə Qazanc/Zərər', kind='total'),
    dict(r=60, label='Yığılmış zərər 2023', kind='loss'), dict(r=61, label='Yığılmış zərər 2024', kind='loss'),
    dict(r=63, label='Mənfəət vergisi', kind='group'),
    dict(r=65, label='Net Qazanc/Zərər', kind='total'),
]


def pl_fold(s):
    """Excel kimi böyük/kiçik hərfə həssas olmayan müqayisə (İ/I/ı -> i)."""
    return str('' if s is None else s).strip().replace('İ', 'i').replace('I', 'i').replace('ı', 'i').lower()


def _pl_num(v):
    return float(v) if isinstance(v, (int, float)) and not isinstance(v, bool) else 0.0


def _pl_int(v):
    try:
        n = float(v)
    except (TypeError, ValueError):
        return None
    return int(n) if n == int(n) else None


def _pl_val(ws, addr):
    return ws[addr].value


def _pl_read_records(ws, s):
    out = []
    last = ws.max_row
    for R in range(s['frm'], last + 1):
        a = _pl_num(ws[f"{s['amt']}{R}"].value)
        label = ws[f"{s['lab']}{R}"].value
        label2 = ws[f"{s['lab2']}{R}"].value if s.get('lab2') else None
        inn = ws[f"{s['inn']}{R}"].value if s.get('inn') else None
        if not a and label is None and inn in (None, ''):
            continue
        if s.get('need_label') and label is None:
            continue
        y = _pl_int(ws[f"{s['y']}{R}"].value) if s.get('y') else None
        m = _pl_int(ws[f"{s['m']}{R}"].value) if s.get('m') else None
        out.append(dict(a=a, label=label, label2=label2, y=y, m=m))
    return out


def _pl_read_payroll(ws):
    out = []
    for R in range(15, ws.max_row + 1):
        y, m = _pl_int(ws[f'B{R}'].value), _pl_int(ws[f'N{R}'].value)
        v = lambda c: _pl_num(ws[f'{c}{R}'].value)
        if not (v('D') or v('E') or v('G') or v('I')):
            continue
        out.append(dict(y=y, m=m, parts={PL_PAY['salary']: v('D'), PL_PAY['dsmf']: v('E'),
                                         PL_PAY['unemployment']: v('I'), PL_PAY['its']: v('G')}))
    return out


def pl_load_workbook(path, progress=None):
    """Excel iş kitabından qeydləri oxuyur. Qaytarır: dict(sheets, pl, rejected)."""
    wb = load_workbook(path, data_only=True)
    sheets, rejected = {}, []
    for key, s in PL_SOURCES.items():
        if s['sheet'] not in wb.sheetnames:
            continue
        ws = wb[s['sheet']]
        if not all(pl_fold(ws[a].value) == pl_fold(w) for a, w in s['head'].items()):
            rejected.append(s['sheet'])
            continue
        if progress:
            progress(f"Oxunur: {s['sheet']}")
        sheets[key] = _pl_read_records(ws, s)
    if 'ə.h' in wb.sheetnames and pl_fold(wb['ə.h']['B2'].value) == pl_fold('Ay'):
        sheets['eh'] = _pl_read_payroll(wb['ə.h'])
    pl = None
    if 'pl' in wb.sheetnames:
        w = wb['pl']
        assets = wb[PL_ASSET_SHEET] if PL_ASSET_SHEET in wb.sheetnames else None
        cells = {}
        for l in PL_LINES:
            cells[l['r']] = {c: w[f"{c}{int(l['r'])}"].value for c in PL_MONTH_COLS + ['AN']
                             if l['r'] == int(l['r']) and isinstance(w[f"{c}{int(l['r'])}"].value, (int, float))}
        pl = dict(year=_pl_int(w['D1'].value),
                  rate=w['C2'].value if isinstance(w['C2'].value, (int, float)) else 0.2,
                  caps=[_pl_num(assets['C1'].value) if assets else 0, _pl_num(assets['C3'].value) if assets else 0],
                  losses=[_pl_num(w['AN60'].value), _pl_num(w['AN61'].value)], cells=cells, extra=0)
    wb.close()
    if pl:  # AN38 düsturundakı "+ rəqəm" (əlavə icarə)
        try:
            wf = load_workbook(path, read_only=True)
            f = next(wf['pl'].iter_rows(min_row=38, max_row=38, min_col=40, max_col=40, values_only=True))[0]
            mm = re.search(r'\+\s*(\d+(?:\.\d+)?)\s*$', str(f or ''))
            pl['extra'] = float(mm.group(1)) if mm else 0
            wf.close()
        except Exception:
            pass
    return dict(sheets=sheets, pl=pl, rejected=rejected)


def pl_build(data):
    """Mənbə qeydlərini 'etiket|il-ay' üzrə toplayır."""
    index = {k: {} for k in list(PL_SOURCES) + ['eh', 'bankK']}
    years = set()

    def add(idx, k, a):
        idx[k] = idx.get(k, 0) + a

    for key, recs in data['sheets'].items():
        s = PL_SOURCES.get(key, {})
        for r in recs:
            if key == 'bank':
                f2 = pl_fold(r['label2'])
                if f2.startswith(('kreditor', 'debitor')) or 'öhdəlik' in f2:
                    continue  # pul hərəkəti, xərc deyil
            if not s.get('spread') and r['y'] and key not in ('gemi', 'ecu'):
                years.add(r['y'])
            if s.get('spread'):
                k = r['y'] if s.get('y') else 'all'
            else:
                k = f"{r['y']}-{r['m']}"
            if key == 'eh':
                for l, v in r['parts'].items():
                    add(index['eh'], f'{pl_fold(l)}|{k}', v)
                continue
            add(index[key], f"{pl_fold(r['label'])}|{k}", r['a'])
            if key == 'bank':
                add(index['bankK'], f"{pl_fold(r['label2'])}|{k}", r['a'])
    return dict(index=index, years=sorted(years), pl=data['pl'])


def _pl_src_amount(ctx, key, label, year, month):
    s = PL_SOURCES.get(key, {})
    idx = ctx['index'].get(key, {})
    f = pl_fold(label)
    if s.get('spread'):
        return idx.get(f"{f}|{year if s.get('y') else 'all'}", 0) / 12
    return idx.get(f'{f}|{year}-{month}', 0)


def pl_compute(ctx, year):
    pl = ctx['pl'] or {}
    rate = pl.get('rate', 0.2)
    caps = pl.get('caps', [0, 0])
    params = (dict(losses=pl['losses'], extra=pl['extra']) if pl and year == pl.get('year')
              else dict(losses=[0, 0], extra=0))
    by = {l['r']: dict(line=l, m=[0.0] * 12, cum=0.0) for l in PL_LINES}
    for l in PL_LINES:
        if not l.get('src'):
            continue
        row = by[l['r']]
        for i in range(12):
            for key in l['src']:
                row['m'][i] += _pl_src_amount(ctx, key, l['label'], year, i + 1)
        row['sum'] = sum(row['m'])
        row['cum'] = row['sum']
        if l.get('cap') is not None and caps[l['cap']]:
            row['capped'] = row['sum'] > caps[l['cap']]
            row['cum'] = min(row['sum'], caps[l['cap']])
        if l.get('extra') and params['extra']:
            row['extra'] = params['extra']
            row['cum'] += params['extra']

    def mix(r, fn):
        by[r]['m'] = [fn(lambda k, i=i: by[k]['m'][i]) for i in range(12)]
        by[r]['cum'] = fn(lambda k: by[k]['cum'])

    for l in PL_LINES:
        if l.get('of'):
            mix(l['r'], lambda g, of=l['of']: sum(g(k) for k in of))
    mix(20, lambda g: g(6) - g(14))
    mix(58, lambda g: g(20) - g(22) - g(55) - g(46))
    by[63]['m'] = [v * rate if v > 0 else 0 for v in by[58]['m']]
    l1, l2 = params['losses']
    by[60]['cum'], by[61]['cum'] = l1, l2
    taxable = by[58]['cum'] - l1 - l2
    by[63]['cum'] = max(taxable, 0) * rate
    by[65]['m'] = [v - by[63]['m'][i] for i, v in enumerate(by[58]['m'])]
    by[65]['cum'] = taxable - by[63]['cum']
    return by


def pl_reconcile(ctx):
    """Yüklənən məlumatı Excel-in 'pl' vərəqində saxlanan dəyərlərlə müqayisə edir."""
    pl = ctx['pl']
    if not pl or not pl.get('year'):
        return None
    by = pl_compute(ctx, pl['year'])
    diffs, checked = [], 0
    for r, row in by.items():
        cells = pl['cells'].get(r) or {}
        for i, col in enumerate(PL_MONTH_COLS + ['AN']):
            if col not in cells:
                continue
            checked += 1
            ours = lambda rr: by[rr]['cum'] if i == 12 else by[rr]['m'][i]
            val = ours(24) + ours(24.5) if r == 24 else ours(r)
            if abs(cells[col] - val) > 0.01:
                diffs.append((r, row['line']['label'], col, cells[col], val))
    return dict(checked=checked, diffs=diffs, year=pl['year'])


def pl_default_year(ctx):
    return (ctx["pl"] or {}).get("year") or (ctx["years"][-1] if ctx["years"] else dt.date.today().year)


def pl_export(path, ctx, year):
    by = pl_compute(ctx, year)
    wb = Workbook()
    ws = wb.active
    ws.title = f"Mənfəət-zərər {year}"
    headers = ["Maddə"] + PL_MONTHS + ["İllik"]
    ws.append(headers)
    _style_header(ws, len(headers))
    for l in PL_LINES:
        r = by[l["r"]]
        ws.append([l["label"]] + [round(x, 2) for x in r["m"]] + [round(r["cum"], 2)])
        bold = l["kind"] != "item"
        for c in range(1, len(headers) + 1):
            cell = ws.cell(row=ws.max_row, column=c)
            cell.font = Font(name=FONT, bold=bold)
            cell.border = BORDER
            if c > 1:
                cell.number_format = MONEY_FMT
            if bold:
                cell.fill = TOTAL_FILL
    ws.column_dimensions["A"].width = 48
    for c in range(2, len(headers) + 1):
        ws.column_dimensions[get_column_letter(c)].width = 15
    ws.freeze_panes = "B2"
    wb.save(path)


# ----------------------------------------------------------------------------
# Chrome uzantısı "Vergi alətləri" v1.6 (qaimə ixracı, ƏDV Əlavə 1, VÖEN yoxla, baza -> proqrama göndər)
# ----------------------------------------------------------------------------
EXT_ZIP_B64 = """
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
UEsFBgAAAAARABEAfAUAAEC2AAAAAA==
"""


def extract_extension(dest):
    """Uzantı fayllarını qovluğa çıxarır: Chrome -> chrome://extensions -> Load unpacked."""
    import zipfile
    dest = Path(dest)
    dest.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(io.BytesIO(base64.b64decode("".join(EXT_ZIP_B64.split())))) as z:
        z.extractall(dest)
    sub = dest / "vv"
    return sub if sub.exists() else dest



# ----------------------------------------------------------------------------
# M.A.Z.A modulu: mapping (təsnifat), layihə, ehtiyat, qeyri-rezident, VÖEN üzrə uzlaşma,
# bazadan Mənfəət-zərər, Pul axını, Balans, Ratio, Kapital.
# Hər şey bu faylın içindədir; əlavə proqram lazım deyil. Excel/xlsb/portal exportu oxunur.
# ----------------------------------------------------------------------------
import zipfile
import struct
import calendar

MZ_SCHEMA = """
CREATE TABLE IF NOT EXISTS projects(id INTEGER PRIMARY KEY, name TEXT UNIQUE, note TEXT);
CREATE TABLE IF NOT EXISTS map_class(id INTEGER PRIMARY KEY, mtype TEXT, pattern TEXT, side TEXT DEFAULT '*',
    cls TEXT, pl_line TEXT, project TEXT, prio INTEGER DEFAULT 5);
CREATE TABLE IF NOT EXISTS map_bank(id INTEGER PRIMARY KEY, mtype TEXT, pattern TEXT, kind TEXT, cf_section TEXT,
    pl_line TEXT, project TEXT, prio INTEGER DEFAULT 5);
CREATE TABLE IF NOT EXISTS map_item(id INTEGER PRIMARY KEY, raw TEXT UNIQUE, std TEXT, unit TEXT);
CREATE TABLE IF NOT EXISTS inv(id INTEGER PRIMARY KEY, src TEXT, dir TEXT, date TEXT, voen TEXT, name TEXT, doc_no TEXT,
    cat TEXT, subcat TEXT, item_code TEXT, item TEXT, item_type TEXT, unit TEXT, qty REAL, base INTEGER, vat INTEGER,
    total INTEGER, status TEXT, pfile TEXT, cls TEXT, pl_line TEXT, project TEXT, rule INTEGER);
CREATE INDEX IF NOT EXISTS ix_inv_voen ON inv(voen, date);
CREATE INDEX IF NOT EXISTS ix_inv_src ON inv(src);
CREATE TABLE IF NOT EXISTS bank(id INTEGER PRIMARY KEY, src TEXT, date TEXT, details TEXT, ref TEXT, voen TEXT,
    inflow INTEGER, outflow INTEGER, balance INTEGER, cat TEXT, sub TEXT, account TEXT, ccy TEXT,
    kind TEXT, cf_section TEXT, pl_line TEXT, project TEXT);
CREATE INDEX IF NOT EXISTS ix_bank_voen ON bank(voen, date);
CREATE INDEX IF NOT EXISTS ix_bank_src ON bank(src);
CREATE TABLE IF NOT EXISTS nonres(id INTEGER PRIMARY KEY, src TEXT, kind TEXT, company TEXT, doc_no TEXT, doc_date TEXT,
    purpose TEXT, category TEXT, project TEXT, ccy TEXT, amount_ccy REAL, amount INTEGER, vat INTEGER, cls TEXT,
    pl_line TEXT, item TEXT, hs TEXT, gb TEXT, note TEXT, manual INTEGER DEFAULT 0);
CREATE TABLE IF NOT EXISTS assets(id INTEGER PRIMARY KEY, name TEXT, category TEXT, cost INTEGER, acc_amort INTEGER,
    rate REAL, start_date TEXT, project TEXT);
CREATE TABLE IF NOT EXISTS payroll(id INTEGER PRIMARY KEY, src TEXT, date TEXT, name TEXT, days INTEGER, salary INTEGER,
    dsmf_e INTEGER, unemp_e INTEGER, its_e INTEGER, tax INTEGER, project TEXT);
CREATE INDEX IF NOT EXISTS ix_pay_src ON payroll(src);
CREATE TABLE IF NOT EXISTS equity(id INTEGER PRIMARY KEY, date TEXT, kind TEXT, amount INTEGER, note TEXT);
CREATE TABLE IF NOT EXISTS opening(id INTEGER PRIMARY KEY, acct TEXT UNIQUE, amount INTEGER);
CREATE TABLE IF NOT EXISTS mz_meta(k TEXT PRIMARY KEY, v TEXT);
"""

MZ_CLS = ["gelir", "maya", "xerc", "ehtiyat", "mal", "esas_vesait", "diger"]
MZ_CLS_LABEL = {"gelir": "Gəlir", "maya": "Maya dəyəri", "xerc": "Əməliyyat xərci", "ehtiyat": "Ehtiyat",
                "mal": "Mal", "esas_vesait": "Əsas vəsait", "diger": "Digər (P&L-ə düşmür)", "": "Təsnifatsız"}
MZ_KINDS = ["Debitor ödəmə", "Kreditor ödəmə", "Bank xərci", "Əmək haqqı", "Vergi və sosial", "Gömrük",
            "Kredit əsas məbləğ", "Kredit faiz", "Təsisçi", "Hesablararası", "İnzibati xərc", "Digər öhdəlik",
            "Təsnifatsız"]
MZ_CF = ["Əməliyyat", "İnvestisiya", "Maliyyə", "Daxili"]
MZ_EQ_KINDS = ["Nizamnamə kapitalı", "Əlavə kapital", "Təsisçi borcu (daxil olan)", "Təsisçi borcu (qaytarılan)", "Dividend"]
MZ_ASSET_CATS = ["Bina tikili və qurğular", "Maşın və avadanlıqılar", "Nəqliyyat vasitələri", "Qeyri maddi aktivlər",
                 "Digər əsas vəsaitlər", "Yüksək texnologiya məhsulları"]
MZ_OPEN_ACCTS = ["Pul", "Debitor", "Kreditor", "Ehtiyat", "Yığılmış mənfəət", "Digər aktiv", "Digər öhdəlik"]
MZ_EXCLUDED_STATUS = ("legv", "cancelled", "canceled", "rejected", "imtina", "etiraz", "deleted", "silin")

# bank növü: (açar söz norm(kateqoriya/təyinat) -> növ, CF bölməsi, P&L sətri istifadə olunsun?)
MZ_BANK_SEED = [
    ("cat", "debitor", "Debitor ödəmə", "Əməliyyat"), ("cat", "kreditor", "Kreditor ödəmə", "Əməliyyat"),
    ("cat", "bank xidmet", "Bank xərci", "Əməliyyat"), ("cat", "emek haqqi", "Əmək haqqı", "Əməliyyat"),
    ("cat", "vergi", "Vergi və sosial", "Əməliyyat"), ("cat", "dsmf", "Vergi və sosial", "Əməliyyat"),
    ("cat", "mdss", "Vergi və sosial", "Əməliyyat"), ("cat", "its", "Vergi və sosial", "Əməliyyat"),
    ("cat", "issizlik", "Vergi və sosial", "Əməliyyat"), ("cat", "gomruk", "Gömrük", "Əməliyyat"),
    ("cat", "kredit esas", "Kredit əsas məbləğ", "Maliyyə"), ("cat", "kredit faiz", "Kredit faiz", "Maliyyə"),
    ("cat", "tesisci", "Təsisçi", "Maliyyə"), ("cat", "hesablararasi", "Hesablararası", "Daxili"),
    ("cat", "inzibati", "İnzibati xərc", "Əməliyyat"), ("cat", "ohdelik", "Digər öhdəlik", "Əməliyyat"),
    ("text", "commission", "Bank xərci", "Əməliyyat"), ("text", "charge for", "Bank xərci", "Əməliyyat"),
    ("text", "komissiya", "Bank xərci", "Əməliyyat"), ("text", "emek haqqi", "Əmək haqqı", "Əməliyyat"),
    ("text", "salary", "Əmək haqqı", "Əməliyyat"), ("text", "convertation", "Hesablararası", "Daxili"),
    ("text", "konvertasiya", "Hesablararası", "Daxili"),
]
MZ_NO_PL_KINDS = ("Debitor ödəmə", "Kreditor ödəmə", "Hesablararası", "Kredit əsas məbləğ", "Təsisçi", "Digər öhdəlik")


def mz_connect(path):
    con = connect(path)
    con.executescript(MZ_SCHEMA)
    return con


def _q(v):
    """AZN -> qəpik (boş/yanlış = 0)."""
    if v is None or isinstance(v, bool):
        return 0
    try:
        return parse_amount(v)
    except Exception:
        return 0


def _az(q):
    return round((q or 0) / 100.0, 2)


def _s(v):
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return str(v).strip()


def _fl(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return 0.0


def _mz_date(v):
    d = parse_date(v)
    if d:
        return d
    s = _s(v)
    if len(s) >= 10 and s[4:5] == "-" and s[7:8] == "-":
        return s[:10]
    return None


def _mz_meta_get(con, k, default=""):
    r = con.execute("SELECT v FROM mz_meta WHERE k=?", (k,)).fetchone()
    return r[0] if r else default


def _mz_meta_set(con, k, v):
    con.execute("INSERT OR REPLACE INTO mz_meta VALUES(?,?)", (k, str(v)))


def mz_settings(cfg=None):
    s = {"rate": 0.2, "loss1": 0.0, "loss2": 0.0, "cap1": 0.0, "cap2": 0.0, "own_voen": ""}
    s.update((cfg or load_cfg()).get("maza", {}))
    return s


# ---------------------------------------------------------------- xlsb oxuyucu (əlavə kitabxana tələb etmir)
def _xb_records(data):
    i, n = 0, len(data)
    while i < n:
        b = data[i]
        i += 1
        rid = b
        if b & 0x80:
            rid = (b & 0x7F) | (data[i] << 7)
            i += 1
        ln, sh = 0, 0
        for _ in range(4):
            b = data[i]
            i += 1
            ln |= (b & 0x7F) << sh
            sh += 7
            if not b & 0x80:
                break
        yield rid, data[i:i + ln]
        i += ln


def _xb_wstr(p, off):
    n = struct.unpack_from("<I", p, off)[0]
    return p[off + 4:off + 4 + 2 * n].decode("utf-16le", "replace"), off + 4 + 2 * n


def _xb_rk(v):
    num = (v >> 2) if v & 2 else struct.unpack("<d", struct.pack("<Q", (v & 0xFFFFFFFC) << 32))[0]
    if v & 2 and num >= 1 << 29:
        num -= 1 << 30
    return num / 100.0 if v & 1 else num


def xlsb_sheets(path):
    """[(vərəq adı, sətirlər generatoru)] - hər sətir dəyərlər siyahısıdır."""
    z = zipfile.ZipFile(path)
    sst = []
    if "xl/sharedStrings.bin" in z.namelist():
        for rid, p in _xb_records(z.read("xl/sharedStrings.bin")):
            if rid == 19:
                sst.append(_xb_wstr(p, 1)[0])
    out = []
    names = sorted((n for n in z.namelist() if re.match(r"xl/worksheets/sheet\d+\.bin$", n)),
                   key=lambda n: int(re.search(r"(\d+)\.bin$", n).group(1)))

    def rows_of(name):
        row, cur, rn = {}, -1, -1
        for rid, p in _xb_records(z.read(name)):
            if rid == 0:
                if row:
                    yield [row.get(c) for c in range(max(row) + 1)]
                row, rn = {}, struct.unpack_from("<I", p, 0)[0]
            elif rid in (1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11):
                col = struct.unpack_from("<I", p, 0)[0]
                if rid == 2:
                    row[col] = _xb_rk(struct.unpack_from("<I", p, 8)[0])
                elif rid == 4 or rid == 10:
                    row[col] = bool(p[8])
                elif rid == 5 or rid == 9:
                    row[col] = struct.unpack_from("<d", p, 8)[0]
                elif rid == 6 or rid == 8:
                    row[col] = _xb_wstr(p, 8)[0]
                elif rid == 7:
                    k = struct.unpack_from("<I", p, 8)[0]
                    row[col] = sst[k] if k < len(sst) else None
        if row:
            yield [row.get(c) for c in range(max(row) + 1)]

    for n in names:
        out.append((n.rsplit("/", 1)[-1][:-4], rows_of(n)))
    return out


def mz_sheets(path):
    """Fayldan [(ad, sətir iteratoru)] (xlsx/xlsm/xlsb)."""
    p = Path(path)
    if p.suffix.lower() == ".xlsb":
        return xlsb_sheets(p)
    wb = load_workbook(p, read_only=True, data_only=True)
    return [(ws.title, ws.iter_rows(values_only=True, max_col=40)) for ws in wb.worksheets]


# ---------------------------------------------------------------- başlıq tanıma
MZ_SPECS = {
    "portal": dict(req=["recv", "send", "no", "item", "total"], map={
        "recv": ["Qəbul edənin VÖEN-i"], "send": ["Göndərənin VÖEN-i"], "no": ["Qaimə nömrəsi"],
        "code": ["Malın (işin, xidmətin) kodu"], "item": ["Malın (işin, xidmətin) adı"],
        "date": ["Yaradılma tarixi", "Qaimə tarixi"], "status": ["Statusu", "Status"],
        "qty": ["Miqdarı, həcmi"], "net": ["Cəmi məbləği"], "vat": ["ƏDV məbləği"], "total": ["Yekun məbləğ"],
        "note": ["Əsas məlumat"], "cat": ["Kateqoriya"], "subcat": ["Subkateqoriya"], "project": ["Proyekt", "Layihə"],
        "partner_name": ["Qarşı tərəfin adı", "Adı"]}),
    "gelen": dict(req=["voen", "date", "item", "base"], map={
        "voen": ["VÖEN"], "name": ["Adı"], "status": ["Vəziyyəti"], "cat": ["Kateqoriya"], "subcat": ["Subkateqoriya"],
        "date": ["Qaimə tarixi"], "ser": ["Qaimə seriyası"], "no": ["Qaimə nömrəsi"], "code": ["Malın (işin, xidmətin) kodu"],
        "item": ["Malın (işin, xidmətin) adı"], "itype": ["Malın (işin, xidmətin) növü"], "unit": ["Ölçü vahidi"],
        "qty": ["Miqdarı, həcmi"], "base": ["Əsas məbləğ", "Malın ƏDV-siz ümumi dəyəri"], "vat": ["Malın ƏDV məbləği"],
        "road": ["Yol vergisi"], "project": ["Layihə", "Proyekt"]}),
    "geden": dict(req=["voen", "date", "base", "total"], map={
        "voen": ["VÖEN"], "name": ["Adı"], "status": ["Vəziyyəti"], "cat": ["Kateqoriya"], "date": ["Qaimə tarixi"],
        "ser": ["Qaimə seriyası"], "no": ["Qaimə nömrəsi"], "base": ["Malın ƏDV-siz ümumi dəyəri"],
        "vat": ["Malın ƏDV məbləği"], "total": ["Yekun məbləğ"], "project": ["Layihə", "Proyekt"]}),
    "ecustom": dict(req=["gb", "date"], map={
        "gb": ["GB sorğu nömrəsi"], "date": ["Bəyannamə tarixi"], "partner": ["Xarici tərəfdaş"], "goods": ["Malın adı"],
        "hs": ["Malın kodu"], "azn": ["Toplam azn dəyər"], "typ": ["növ"], "asset": ["Əsas vəsaitin adı"],
        "amount": ["Əsas məbləğ"], "vat": ["ƏDV"], "project": ["Layihə", "Proyekt"]}),
    "bank": dict(req=["date", "details", "out"], map={
        "date": ["Əməliyyat tarixi"], "date2": ["İcra tarixi"], "details": ["Təyinat"], "ref": ["İstinad No"],
        "voen": ["VÖEN"], "inn": ["Mədaxil"], "out": ["Məxaric"], "bal": ["Balans"], "cat": ["Ümumi kateqoriya"],
        "sub": ["növ"], "acct": ["Bank hesabı"], "ccy": ["Valyuta növ"], "project": ["Layihə", "Proyekt"]}),
    "gemi": dict(req=["company", "inv_no", "inv_date", "total"], map={
        "company": ["Şirkətin adı"], "inv_no": ["İnvoice №"], "inv_date": ["İnvoice tarixi"], "ship": ["Gəmi"],
        "purpose": ["Təyinat"], "cat": ["Təsnifat"], "ccy": ["Valyuta"], "total": ["Yekun"], "azn": ["AZN"],
        "project": ["Layihə", "Proyekt"]}),
    "payroll": dict(req=["date", "salary", "dsmf_e"], map={
        "date": ["Tarix"], "name": ["Ad, soyad, ata adı"], "days": ["İşlədiyi gün"], "salary": ["Əmək haqqı"],
        "tax": ["Gəlir vergisi"], "dsmf_e": ["DSMF — işəgötürən", "DSMF - işəgötürən"],
        "unemp_e": ["İşsizlik — işəgötürən", "İşsizlik - işəgötürən"], "its_e": ["İTS — işəgötürən", "İTS - işəgötürən"],
        "project": ["Layihə", "Proyekt"]}),
}
MZ_ORDER = ["portal", "gelen", "geden", "ecustom", "bank", "gemi", "payroll"]


def _mz_find_header(rows_buf, spec):
    for hi, row in enumerate(rows_buf[:15]):
        cells = [norm(c) if c is not None else "" for c in row]
        m = {}
        for key, alts in spec["map"].items():
            for a in alts:
                na = norm(a)
                if na in cells:
                    m[key] = cells.index(na)
                    break
            else:
                for a in alts:
                    na = norm(a)
                    if len(na) > 4:
                        ix = [i for i, c in enumerate(cells) if c and na in c]
                        if ix:
                            m[key] = ix[0]
                            break
        if all(k in m for k in spec["req"]):
            return hi, m
    return None, None


def _mz_detect(rows_buf):
    for kind in MZ_ORDER:
        hi, m = _mz_find_header(rows_buf, MZ_SPECS[kind])
        if hi is not None:
            return kind, hi, m
    return None, None, None


def _mz_flag_cancelled(status):
    n = norm(status or "")
    return bool(n) and "refused" not in n and any(w in n for w in MZ_EXCLUDED_STATUS)


# ---------------------------------------------------------------- import
def mz_import_path(con, path, progress=None):
    """Fayl və ya qovluq. Hər vərəq başlığına görə tanınır. Qaytarır: [(fayl|vərəq, növ, sətir, dublikat)]."""
    p = Path(str(path).strip().strip('"'))
    if p.is_dir():
        files = sorted(f for f in p.rglob("*") if f.suffix.lower() in (".xlsx", ".xlsm", ".xlsb")
                       and not f.name.startswith("~$"))
    elif p.exists():
        files = [p]
    else:
        raise ValueError("Fayl/qovluq tapılmadı: " + str(p))
    own = mz_settings().get("own_voen") or load_cfg().get("company", "")
    report = []
    for f in files:
        try:
            sheets = mz_sheets(f)
        except Exception as e:
            report.append((f.name, "açılmadı: " + str(e), 0, 0))
            continue
        for sname, it in sheets:
            if norm(sname) == "telimat":      # şablonun izahat vərəqi
                continue
            buf = []
            gen = iter(it)
            for r in gen:
                buf.append(list(r))
                if len(buf) >= 15:
                    break
            kind, hi, m = _mz_detect(buf)
            if not kind:
                continue
            if progress:
                progress(f"Oxunur: {f.name} / {sname} ({kind})")
            rows = (r for r in (buf[hi + 1:] + [list(x) for x in gen]))
            src = f"{f.name}|{sname}"
            n, dup = mz_load(con, kind, src, rows, m, own)
            report.append((src, kind, n, dup))
    mz_seed(con)
    mz_classify(con)
    return report


def _g(row, m, key):
    i = m.get(key)
    return row[i] if i is not None and i < len(row) else None


def mz_load(con, kind, src, rows, m, own=""):
    """Bir vərəqi bazaya yazır (eyni mənbə əvvəlcə silinir; başqa mənbədə eyni sətir varsa keçilir)."""
    tab = {"portal": "inv", "gelen": "inv", "geden": "inv", "ecustom": "nonres", "gemi": "nonres",
           "bank": "bank", "payroll": "payroll"}[kind]
    con.execute(f"DELETE FROM {tab} WHERE src=?", (src,))
    out, dup, seen = [], 0, set()
    if tab == "inv":
        have = {(a, b, c, d, e) for a, b, c, d, e in con.execute("SELECT dir,voen,doc_no,date,base||'|'||item FROM inv")}
    elif tab == "bank":
        have = {(a, b, c, d, e) for a, b, c, d, e in con.execute("SELECT date,ref,outflow,inflow,account FROM bank")}
    else:
        have = set()
    portal_rows = []
    for row in rows:
        if not row or all(c is None or c == "" for c in row):
            continue
        if kind == "portal":
            portal_rows.append(row)
            continue
        if kind in ("gelen", "geden"):
            d = _mz_date(_g(row, m, "date"))
            voen = parse_voen(_g(row, m, "voen"))
            if not d or not voen:
                continue
            st = _s(_g(row, m, "status"))
            if _mz_flag_cancelled(st):
                continue
            base = _q(_g(row, m, "base"))
            vat = _q(_g(row, m, "vat"))
            total = _q(_g(row, m, "total")) if kind == "geden" else base + vat + _q(_g(row, m, "road"))
            doc = (_s(_g(row, m, "ser")) + " " + _s(_g(row, m, "no"))).strip()
            item = _s(_g(row, m, "item")) if kind == "gelen" else ""
            key = ("G" if kind == "gelen" else "S", voen, doc, d, f"{base}|{item}")
            if key in have:
                dup += 1
                continue
            out.append((src, key[0], d, voen, _s(_g(row, m, "name")), doc, _s(_g(row, m, "cat")),
                        _s(_g(row, m, "subcat")), _s(_g(row, m, "code")), item, _s(_g(row, m, "itype")),
                        _s(_g(row, m, "unit")).lower(), _fl(_g(row, m, "qty")), base, vat, total, st,
                        _s(_g(row, m, "project"))))
        elif kind == "bank":
            d = _mz_date(_g(row, m, "date")) or _mz_date(_g(row, m, "date2"))
            if not d:
                continue
            inn, outv = _q(_g(row, m, "inn")), _q(_g(row, m, "out"))
            det, ref = _s(_g(row, m, "details")), _s(_g(row, m, "ref"))
            key = (d, ref, outv, inn, _s(_g(row, m, "acct")))
            if key in have:
                dup += 1
                continue
            out.append((src, d, det, ref, parse_voen(_g(row, m, "voen")) or "", inn, outv, _q(_g(row, m, "bal")),
                        _s(_g(row, m, "cat")), _s(_g(row, m, "sub")), _s(_g(row, m, "acct")),
                        _s(_g(row, m, "ccy")).lower() or "azn", _s(_g(row, m, "project"))))
        elif kind == "ecustom":
            d = _mz_date(_g(row, m, "date"))
            if not d:
                continue
            amt = _q(_g(row, m, "amount")) or _q(_g(row, m, "azn"))
            typ = _s(_g(row, m, "typ"))
            out.append((src, "idxal", _s(_g(row, m, "partner")), _s(_g(row, m, "gb")), d, _s(_g(row, m, "goods"))[:200],
                        typ, _s(_g(row, m, "project")), "azn", _fl(_g(row, m, "azn")), amt, _q(_g(row, m, "vat")),
                        _s(_g(row, m, "asset")), _s(_g(row, m, "hs")), _s(_g(row, m, "gb"))))
        elif kind == "gemi":
            d = _mz_date(_g(row, m, "inv_date"))
            if not d:
                continue
            amt = _q(_g(row, m, "total")) or _q(_g(row, m, "azn"))
            out.append((src, "xidmet", _s(_g(row, m, "company")), _s(_g(row, m, "inv_no")), d,
                        (_s(_g(row, m, "ship")) + " " + _s(_g(row, m, "purpose"))).strip(), _s(_g(row, m, "cat")),
                        _s(_g(row, m, "project")), (_s(_g(row, m, "ccy")) or "azn").lower(), 0.0, amt, 0, "", "",
                        _s(_g(row, m, "inv_no"))))
        elif kind == "payroll":
            d = _mz_date(_g(row, m, "date"))
            if not d:
                continue
            out.append((src, d, _s(_g(row, m, "name")), int(_fl(_g(row, m, "days"))), _q(_g(row, m, "salary")),
                        _q(_g(row, m, "dsmf_e")), _q(_g(row, m, "unemp_e")), _q(_g(row, m, "its_e")),
                        _q(_g(row, m, "tax")), _s(_g(row, m, "project"))))
    if kind == "portal":
        out, dup = _mz_portal(con, src, portal_rows, m, own, have)
        tab = "inv"
    if not out:
        return 0, dup
    with con:
        if tab == "inv":
            con.executemany("INSERT INTO inv(src,dir,date,voen,name,doc_no,cat,subcat,item_code,item,item_type,unit,qty,"
                            "base,vat,total,status,pfile) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", out)
            con.executemany("INSERT INTO parties(voen,name) VALUES(?,?) ON CONFLICT(voen) DO UPDATE SET "
                            "name=CASE WHEN COALESCE(parties.name,'')='' THEN excluded.name ELSE parties.name END",
                            {(r[3], r[4]) for r in out if r[4]})
        elif tab == "bank":
            con.executemany("INSERT INTO bank(src,date,details,ref,voen,inflow,outflow,balance,cat,sub,account,ccy,project)"
                            " VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)", out)
        elif tab == "nonres":
            con.executemany("INSERT INTO nonres(src,kind,company,doc_no,doc_date,purpose,category,project,ccy,amount_ccy,"
                            "amount,vat,item,hs,gb) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", out)
        else:
            con.executemany("INSERT INTO payroll(src,date,name,days,salary,dsmf_e,unemp_e,its_e,tax,project) "
                            "VALUES(?,?,?,?,?,?,?,?,?,?)", out)
    return len(out), dup


def _mz_portal(con, src, rows, m, own, have):
    """Portal exportu (qəbul edən + göndərən VÖEN). Öz VÖEN-imiz: ayarda yoxdursa ən çox rast gəlinən."""
    if not own:
        cnt = {}
        for r in rows:
            for k in ("recv", "send"):
                v = parse_voen(_g(r, m, k))
                if v:
                    cnt[v] = cnt.get(v, 0) + 1
        own = max(cnt, key=cnt.get) if cnt else ""
    out, dup, seen = [], 0, set()
    for r in rows:
        recv, send = parse_voen(_g(r, m, "recv")), parse_voen(_g(r, m, "send"))
        d = _mz_date(_g(r, m, "date"))
        st = _s(_g(r, m, "status"))
        if not d or not (recv and send) or _mz_flag_cancelled(st):
            continue
        direction = "S" if send == own else "G"
        voen = recv if direction == "S" else send
        total = _q(_g(r, m, "total"))
        vat = _q(_g(r, m, "vat"))
        base = _q(_g(r, m, "net")) or (total - vat)
        item = _s(_g(r, m, "item"))
        doc = _s(_g(r, m, "no"))
        key = (direction, voen, doc, d, f"{base}|{item}")
        if key in have:
            dup += 1
            continue
        out.append((src, direction, d, voen, _s(_g(r, m, "partner_name")), doc, _s(_g(r, m, "cat")),
                    _s(_g(r, m, "subcat")), _s(_g(r, m, "code")), item, "", "", _fl(_g(r, m, "qty")), base, vat,
                    total, st, _s(_g(r, m, "project"))))
    return out, dup


# ---------------------------------------------------------------- mapping seed və təsnifat
def _pl_label_class():
    """P&L maddəsinin adı -> təsnif sinfi (gelir/maya/xerc)."""
    out = {}
    for l in PL_LINES:
        if l["kind"] != "item" or not l.get("src"):
            continue
        r = int(l["r"])
        cls = "gelir" if 7 <= r <= 12 else "maya" if 15 <= r <= 17 else "xerc" if 23 <= r <= 44 else None
        if cls and not (l["src"] == ["eh"]):
            out[pl_fold(l["label"])] = (cls, l["label"])
    return out


def mz_pl_labels():
    return [l["label"] for l in PL_LINES if l["kind"] == "item" and l["r"] != 24.5 and not (47 <= l["r"] <= 52)]


def mz_seed(con):
    """Mövcud datadakı kateqoriyalardan mapping qaydaları yaradır (yalnız olmayanları)."""
    lab = _pl_label_class()
    have = {(r[0], pl_fold(r[1])) for r in con.execute("SELECT mtype,pattern FROM map_class")}
    new = []
    pats = {r[0] for r in con.execute("SELECT DISTINCT COALESCE(NULLIF(subcat,''),cat) FROM inv")}
    pats |= {r[0] for r in con.execute("SELECT DISTINCT category FROM nonres")}
    for pat in pats:
        if not pat or ("subcat", pl_fold(pat)) in have:
            continue
        f = pl_fold(pat)
        if f in lab:
            new.append(("subcat", pat, "*", lab[f][0], lab[f][1], "", 9))
        elif any(norm(a) == norm(pat) or (len(norm(pat)) > 8 and norm(a).startswith(norm(pat)[:10])) for a in MZ_ASSET_CATS):
            new.append(("subcat", pat, "*", "esas_vesait", "", "", 9))
    if new:
        con.executemany("INSERT INTO map_class(mtype,pattern,side,cls,pl_line,project,prio) VALUES(?,?,?,?,?,?,?)", new)
    if not con.execute("SELECT 1 FROM map_bank LIMIT 1").fetchone():
        con.executemany("INSERT INTO map_bank(mtype,pattern,kind,cf_section,pl_line,project,prio) VALUES(?,?,?,?,?,?,9)",
                        [(t, p, k, c, "Bank xidmət haqqı" if k == "Bank xərci" and t == "cat" else "", "")
                         for t, p, k, c in MZ_BANK_SEED])


def mz_classify(con):
    """Qaimə, qeyri-rezident və bank sətirlərinə qaydaları tətbiq edir."""
    rules = con.execute("SELECT id,mtype,pattern,side,cls,pl_line,project FROM map_class ORDER BY prio,id").fetchall()
    R = [(r[0], r[1], r[2] if r[1] == "voen" else pl_fold(r[2]), r[3] or "*", r[4] or "", r[5] or "", r[6] or "") for r in rules]

    def match(side, voen, name, item, sub):
        for rid, mt, pat, sd, cls, pl, prj in R:
            if sd not in ("*", "", side):
                continue
            if mt == "voen":
                ok = voen == pat
            elif mt == "name":
                ok = pat in pl_fold(name)
            elif mt == "item":
                ok = pat in pl_fold(item)
            else:
                ok = pl_fold(sub) == pat
            if ok:
                return rid, cls, pl, prj
        return None

    upd, hist = [], {}
    rows = con.execute("SELECT id,dir,voen,name,item,COALESCE(NULLIF(subcat,''),cat),pfile FROM inv").fetchall()
    res = {}
    for iid, d, voen, name, item, sub, pf in rows:
        m = match(d, voen, name or "", item or "", sub or "")
        if m:
            res[iid] = m
            h = hist.setdefault((d, voen), {})
            h[m[1:3]] = h.get(m[1:3], 0) + 1
    dom = {}
    for k, h in hist.items():
        tot = sum(h.values())
        best = max(h, key=h.get)
        if h[best] / tot >= 0.8 and best[0]:
            dom[k] = best
    for iid, d, voen, name, item, sub, pf in rows:
        if iid in res:
            rid, cls, pl, prj = res[iid]
        elif (d, voen) in dom:
            rid, (cls, pl), prj = -1, dom[(d, voen)], ""
        else:
            rid, cls, pl, prj = 0, "", "", ""
        upd.append((cls, pl, prj or pf or "", rid, iid))
    with con:
        con.executemany("UPDATE inv SET cls=?,pl_line=?,project=?,rule=? WHERE id=?", upd)
    # qeyri-rezident
    upd = []
    for nid, comp, cat, goods, man in con.execute(
            "SELECT id,company,category,COALESCE(item,'')||' '||COALESCE(purpose,''),manual FROM nonres"):
        if man:
            continue
        m = match("N", "", comp or "", goods or "", cat or "")
        upd.append((m[1], m[2], m[3], nid) if m else ("", "", "", nid))
    with con:
        for cls, pl, prj, nid in upd:
            con.execute("UPDATE nonres SET cls=?,pl_line=?,project=COALESCE(NULLIF(?,''),project) WHERE id=?", (cls, pl, prj, nid))
    mz_classify_bank(con)


def mz_classify_bank(con):
    rules = con.execute("SELECT mtype,pattern,kind,cf_section,pl_line,project FROM map_bank ORDER BY prio,id").fetchall()
    R = [(r[0], r[1] if r[0] == "voen" else re.compile(r"(^|\W)" + re.escape(norm(r[1]))), r[2], r[3], r[4] or "", r[5] or "")
         for r in rules]
    s_voen = {r[0] for r in con.execute("SELECT DISTINCT voen FROM inv WHERE dir='S'")}
    g_voen = {r[0] for r in con.execute("SELECT DISTINCT voen FROM inv WHERE dir='G'")}
    upd = []
    for bid, voen, det, cat, sub, inn, out in con.execute(
            "SELECT id,voen,details,cat,sub,inflow,outflow FROM bank"):
        ndet, ncat = norm(det or ""), norm(cat or "")
        hit = None
        for mt, pat, kind, cf, pl, prj in R:
            if (mt == "voen" and voen == pat) or (mt == "text" and pat.search(ndet)) or (mt == "cat" and pat.search(ncat)):
                hit = (kind, cf, pl, prj)
                break
        if not hit and voen:
            if voen in s_voen and inn > 0:
                hit = ("Debitor ödəmə", "Əməliyyat", "", "")
            elif voen in g_voen and out > 0:
                hit = ("Kreditor ödəmə", "Əməliyyat", "", "")
        if not hit:
            hit = ("Təsnifatsız", "Əməliyyat", "", "")
        kind, cf, pl, prj = hit
        if kind == "Bank xərci" and not pl:
            pl = "Bank xidmət haqqı"
        if not pl and kind not in MZ_NO_PL_KINDS and sub:
            pl = sub                       # köhnə iş kitabındakı qayda: "növ" sütunu = P&L maddəsi
        if kind in MZ_NO_PL_KINDS:
            pl = ""
        upd.append((kind, cf, pl, prj, bid))
    with con:
        con.executemany("UPDATE bank SET kind=?,cf_section=?,pl_line=?,project=? WHERE id=?", upd)


# ---------------------------------------------------------------- əsas vəsait amortizasiyası
def _asset_sched(cost, acc, rate, sd, ty, tm=12):
    """Azalan qalıq üsulu. (ty ilinin sonunda/tm ayına yığılmış amortizasiya, ty ilinin amortizasiyası)."""
    sd = sd or "2000-12-31"
    sy, sm = int(sd[:4]), int(sd[5:7])
    rate = rate or 0
    rate = rate / 100 if rate > 1 else rate
    last = 0.0
    for y in range(sy, ty + 1):
        lo = sm + 1 if y == sy else 1
        hi = tm if y == ty else 12
        months = max(0, hi - lo + 1)
        last = max((cost - acc) * rate * months / 12.0, 0) if months else 0.0
        acc += last
    return acc, (last if sy <= ty else 0.0)


def mz_amort(con, year, project=None):
    """{kateqoriya: ilin amortizasiyası (AZN)}"""
    out = {}
    for name, cat, cost, acc, rate, sd, prj in con.execute("SELECT name,category,cost,acc_amort,rate,start_date,project FROM assets"):
        if project and prj != project:
            continue
        if int((sd or "2000")[:4]) > year:
            continue
        _, a = _asset_sched(cost or 0, acc or 0, rate, sd, year)
        out[cat or ""] = out.get(cat or "", 0) + a / 100.0
    return out


def mz_asset_net(con, as_of):
    ay, am = int(as_of[:4]), int(as_of[5:7])
    gross = accum = 0
    for cost, acc, rate, sd in con.execute("SELECT cost,acc_amort,rate,start_date FROM assets WHERE COALESCE(start_date,'')<=?", (as_of,)):
        gross += cost or 0
        accum += _asset_sched(cost or 0, acc or 0, rate, sd, ay, am)[0]
    return gross, accum


# ---------------------------------------------------------------- Mənfəət-zərər (bazadan)
def mz_pl_ctx(con, year, project=None):
    cfgm = mz_settings()
    idx = {k: {} for k in list(PL_SOURCES) + ["eh", "bankK"]}
    years = set()
    P = project or ""

    def add(k, label, d, amt):
        if not d or len(d) < 7:
            return
        y, mth = int(d[:4]), int(d[5:7])
        years.add(y)
        key = f"{pl_fold(label)}|{y}-{mth}"
        idx[k][key] = idx[k].get(key, 0) + amt

    for dr, d, pl, base in con.execute(
            "SELECT dir,date,pl_line,base FROM inv WHERE pl_line<>'' AND cls IN ('gelir','maya','xerc') "
            "AND (?='' OR project=?)", (P, P)):
        add("ged" if dr == "S" else "gel", pl, d, base / 100.0)
    for kind, d, pl, amt in con.execute(
            "SELECT kind,doc_date,pl_line,amount FROM nonres WHERE pl_line<>'' AND cls IN ('gelir','maya','xerc') "
            "AND (?='' OR project=?)", (P, P)):
        add("ecu" if kind == "idxal" else "gemi", pl, d, amt / 100.0)
    for d, pl, inn, out in con.execute(
            "SELECT date,pl_line,inflow,outflow FROM bank WHERE pl_line<>'' AND (?='' OR project=?)", (P, P)):
        add("bankK" if pl_fold(pl) == pl_fold("Bank xidmət haqqı") else "bank", pl, d, (out - inn) / 100.0)
    for d, sal, ds, un, its, prj in con.execute("SELECT date,salary,dsmf_e,unemp_e,its_e,project FROM payroll"):
        if P and prj != P:
            continue
        for k, v in (("salary", sal), ("dsmf", ds), ("unemployment", un), ("its", its)):
            add("eh", PL_PAY[k], d, (v or 0) / 100.0)
    for cat, a in mz_amort(con, year, project).items():
        idx["ev"][f"{pl_fold(cat + ' Amortizasiya')}|all"] = a
    ctx = dict(index=idx, years=sorted(years or {year}),
               pl=dict(year=year, rate=cfgm["rate"], caps=[cfgm["cap1"], cfgm["cap2"]],
                       losses=[cfgm["loss1"], cfgm["loss2"]], extra=0))
    return ctx


def mz_pl(con, year, project=None):
    ctx = mz_pl_ctx(con, year, project)
    by = pl_compute(ctx, year)
    rows = []
    for l in PL_LINES:
        r = by[l["r"]]
        rows.append({"c": [l["label"]] + [round(x, 2) for x in r["m"]] + [round(r["cum"], 2)],
                     "k": "" if l["kind"] == "item" else "grp"})
    return dict(cols=[["Maddə", 0]] + [[m, 1] for m in PL_MONTHS] + [["İllik", 1]], rows=rows,
                years=ctx["years"], by=by,
                note=f"{year} il" + (f" | layihə: {project}" if project else " | bütün layihələr") +
                     f" | mənfəət vergisi {mz_settings()['rate'] * 100:g}%")


def mz_pl_period(con, d1, d2, project=None):
    """[d1,d2] aralığında P&L sətrləri üzrə cəm {r: AZN}."""
    d1 = d1 or "1900-01-01"
    d2 = d2 or dt.date.today().isoformat()
    y1, y2 = int(d1[:4]), int(d2[:4])
    tot = {l["r"]: 0.0 for l in PL_LINES}
    for y in range(y1, y2 + 1):
        by = pl_compute(mz_pl_ctx(con, y, project), y)
        for mi in range(12):
            first = f"{y}-{mi + 1:02d}-01"
            last = f"{y}-{mi + 1:02d}-{calendar.monthrange(y, mi + 1)[1]:02d}"
            if last < d1 or first > d2:
                continue
            for r, v in by.items():
                tot[r] += v["m"][mi]
    return tot


# ---------------------------------------------------------------- VÖEN üzrə uzlaşma
def mz_voen(con, d1, d2, project=None):
    d1 = d1 or "1900-01-01"
    d2 = d2 or "9999-12-31"
    P = project or ""
    A = {}

    def acc(v, key, amt, name=""):
        e = A.setdefault(v, dict(name="", s=0, dp=0, sb=0, g=0, kp=0, gb=0, other=0))
        if name and not e["name"]:
            e["name"] = name
        e[key] += amt

    for dr, voen, name, d, tot in con.execute("SELECT dir,voen,name,date,total FROM inv WHERE date<=? AND (?='' OR project=?)", (d2, P, P)):
        if dr == "S":
            acc(voen, "sb", tot, name)
            if d >= d1:
                acc(voen, "s", tot)
        else:
            acc(voen, "gb", tot, name)
            if d >= d1:
                acc(voen, "g", tot)
    for voen, kind, d, inn, out in con.execute("SELECT voen,kind,date,inflow,outflow FROM bank WHERE voen<>'' AND date<=? "
                                               "AND (?='' OR project=?)", (d2, P, P)):
        if kind == "Debitor ödəmə":
            acc(voen, "sb", -(inn - out))
            if d >= d1:
                acc(voen, "dp", inn - out)
        elif kind == "Kreditor ödəmə":
            acc(voen, "gb", -(out - inn))
            if d >= d1:
                acc(voen, "kp", out - inn)
        elif d >= d1 and kind not in ("Hesablararası",):
            acc(voen, "other", out - inn)
    names = dict(con.execute("SELECT voen, COALESCE(NULLIF(reg_name,''),name,'') FROM parties"))
    rows = []
    for v, e in sorted(A.items(), key=lambda x: -abs(x[1]["sb"] - x[1]["gb"])):
        deb, kred = e["sb"], e["gb"]
        rows.append({"c": [v, e["name"] or names.get(v, ""), _az(e["s"]), _az(e["dp"]), _az(deb), _az(e["g"]),
                           _az(e["kp"]), _az(kred), _az(e["other"]), _az(deb - kred)], "k": ""})
    tot = ["", "YEKUN"] + [round(sum(r["c"][i] for r in rows), 2) for i in range(2, 10)]
    rows.append({"c": tot, "k": "tot"})
    return dict(cols=[["VÖEN", 0], ["Tərəf", 0], ["Gedən qaimə (dövr)", 1], ["Debitor ödəmə (dövr)", 1],
                      ["Debitor qalıq", 1], ["Gələn qaimə (dövr)", 1], ["Kreditor ödəmə (dövr)", 1],
                      ["Kreditor qalıq", 1], ["Digər bank hərəkəti", 1], ["Xalis (D − K)", 1]], rows=rows,
                note="Qalıqlar d2 tarixinə qədərki bütün sənədlərlədir. Mənfi debitor = alınmış avans; mənfi kreditor = verilmiş avans.")


def mz_net_positions(con, as_of):
    """(debitor, kreditor) qəpiklə. Hər VÖEN üzrə: müsbət debitor = aktiv, mənfi = alınmış avans (öhdəlik);
    müsbət kreditor = öhdəlik, mənfi = verilmiş avans (aktiv)."""
    P = {}
    for dr, voen, tot in con.execute("SELECT dir,voen,total FROM inv WHERE date<=?", (as_of,)):
        e = P.setdefault(voen, [0, 0])
        e[0 if dr == "S" else 1] += tot
    for voen, kind, inn, out in con.execute("SELECT voen,kind,inflow,outflow FROM bank WHERE voen<>'' AND date<=?", (as_of,)):
        e = P.setdefault(voen, [0, 0])
        if kind == "Debitor ödəmə":
            e[0] -= inn - out
        elif kind == "Kreditor ödəmə":
            e[1] -= out - inn
    deb = sum(max(d, 0) + max(-k, 0) for d, k in P.values())
    kred = sum(max(k, 0) + max(-d, 0) for d, k in P.values())
    return deb, kred


# ---------------------------------------------------------------- ehtiyat
def _mz_std(con):
    return {pl_fold(a): (b or a) for a, b in con.execute("SELECT raw,std FROM map_item")}


def mz_stock(con, d1, d2, project=None):
    d1 = d1 or "1900-01-01"
    d2 = d2 or "9999-12-31"
    P = project or ""
    std = _mz_std(con)
    items = {}
    stocked = {std.get(pl_fold(i), i) for (i,) in con.execute(
        "SELECT item FROM inv WHERE dir='G' AND cls IN ('ehtiyat','mal') AND item<>''")}
    for dr, d, item, unit, qty, base in con.execute(
            "SELECT dir,date,item,unit,qty,base FROM inv WHERE item<>'' AND date<=? AND (?='' OR project=?) "
            "AND ((dir='G' AND cls IN ('ehtiyat','mal')) OR dir='S') ORDER BY date,id", (d2, P, P)):
        key = std.get(pl_fold(item), item)
        if key not in stocked:
            continue
        e = items.setdefault(key, dict(unit=unit, q=0.0, v=0.0, oq=0.0, ov=0.0, iq=0.0, iv=0.0, xq=0.0, xv=0.0, neg=False))
        if not e["unit"]:
            e["unit"] = unit
        pre = d < d1
        if dr == "G":
            e["q"] += qty
            e["v"] += base
            if pre:
                e["oq"] += qty
                e["ov"] += base
            else:
                e["iq"] += qty
                e["iv"] += base
        else:
            avg = e["v"] / e["q"] if e["q"] > 0 else 0
            cost = avg * qty
            e["q"] -= qty
            e["v"] -= cost
            if e["q"] < -1e-9:
                e["neg"] = True
            if pre:
                e["oq"] -= qty
                e["ov"] -= cost
            else:
                e["xq"] += qty
                e["xv"] += cost
    rows = [{"c": [k, e["unit"], round(e["oq"], 3), _az(e["ov"]), round(e["iq"], 3), _az(e["iv"]), round(e["xq"], 3),
                   _az(e["xv"]), round(e["q"], 3), _az(e["v"]), "mənfi qalıq!" if e["neg"] else ""],
             "k": "risk" if e["neg"] else ""} for k, e in sorted(items.items())]
    rows.append({"c": ["YEKUN", "", "", _az(sum(e["ov"] for e in items.values())), "", _az(sum(e["iv"] for e in items.values())),
                       "", _az(sum(e["xv"] for e in items.values())), "", _az(sum(e["v"] for e in items.values())), ""],
                 "k": "tot"})
    return dict(cols=[["Mal / ehtiyat", 0], ["Vahid", 0], ["Əvv. miqdar", 1], ["Əvv. məbləğ", 1], ["Daxilolma miqdar", 1],
                      ["Daxilolma məbləğ", 1], ["Çıxış miqdar", 1], ["Çıxış (orta maya)", 1], ["Son miqdar", 1],
                      ["Son məbləğ", 1], ["Qeyd", 0]], rows=rows,
                note="Sinfi 'ehtiyat' və ya 'mal' olan gələn qaimələr daxilolmadır; eyni adlı (mapping: ad uyğunlaşdırma) "
                     "gedən qaimə sətirləri orta maya ilə çıxışdır.", total=sum(e["v"] for e in items.values()))


# ---------------------------------------------------------------- pul axını
def mz_cf(con, year, project=None):
    P = project or ""
    mon = [f"{year}-{m:02d}" for m in range(1, 13)]
    M = {}
    for d, kind, cf, inn, out in con.execute("SELECT date,kind,cf_section,inflow,outflow FROM bank WHERE substr(date,1,4)=? "
                                             "AND (?='' OR project=?)", (str(year), P, P)):
        e = M.setdefault((cf or "Əməliyyat", kind or "Təsnifatsız"), [0] * 12)
        e[int(d[5:7]) - 1] += inn - out
    rows, grand = [], [0] * 12
    for sec in MZ_CF:
        ks = sorted(k for s, k in M if s == sec)
        if not ks:
            continue
        sub = [0] * 12
        rows.append({"c": [sec] + [""] * 13, "k": "grp"})
        for k in ks:
            v = M[(sec, k)]
            rows.append({"c": ["   " + k] + [_az(x) for x in v] + [_az(sum(v))], "k": ""})
            sub = [a + b for a, b in zip(sub, v)]
        rows.append({"c": [f"{sec} üzrə xalis"] + [_az(x) for x in sub] + [_az(sum(sub))], "k": "tot"})
        if sec != "Daxili":
            grand = [a + b for a, b in zip(grand, sub)]
    op = _mz_cash_before(con, f"{year}-01-01", P)
    rows.append({"c": ["Xalis pul axını (daxili köçürmələr istisna)"] + [_az(x) for x in grand] + [_az(sum(grand))], "k": "grp"})
    run, opens, closes = op, [], []
    for x in grand:
        opens.append(run)
        run += x
        closes.append(run)
    rows.append({"c": ["Dövrün əvvəlinə pul"] + [_az(x) for x in opens] + [_az(op)], "k": ""})
    rows.append({"c": ["Dövrün sonuna pul"] + [_az(x) for x in closes] + [_az(run)], "k": "tot"})
    return dict(cols=[["Maddə", 0]] + [[m, 1] for m in PL_MONTHS] + [["Cəmi", 1]], rows=rows,
                note=f"{year} | bank çıxarışından (mədaxil − məxaric); növ və bölmə Mapping > Bank qaydalarından.")


def _mz_cash_before(con, d, P=""):
    r = con.execute("SELECT COALESCE(SUM(inflow-outflow),0) FROM bank WHERE date<? AND (?='' OR project=?)", (d, P, P)).fetchone()[0]
    o = con.execute("SELECT amount FROM opening WHERE acct='Pul'").fetchone()
    return r + (o[0] if o and not P else 0)


# ---------------------------------------------------------------- balans
def mz_equity_totals(con, as_of):
    t = {k: 0 for k in MZ_EQ_KINDS}
    for k, a in con.execute("SELECT kind,amount FROM equity WHERE date<=?", (as_of,)):
        if k in t:
            t[k] += a
    return t


def mz_bs_data(con, as_of, project=None):
    P = project or ""
    cash = _mz_cash_before(con, (dt.date.fromisoformat(as_of) + dt.timedelta(days=1)).isoformat(), P)
    deb, kred = mz_net_positions(con, as_of)
    op = dict(con.execute("SELECT acct,amount FROM opening"))
    stock = mz_stock(con, None, as_of, project)["total"] + op.get("Ehtiyat", 0)
    gross, acc = mz_asset_net(con, as_of)
    eq = mz_equity_totals(con, as_of)
    profit = 0.0
    for y in range(int(min([r[0][:4] for r in con.execute("SELECT MIN(date) FROM bank UNION SELECT MIN(date) FROM inv") if r[0]] or [as_of[:4]])),
                   int(as_of[:4]) + 1):
        ctx = mz_pl_ctx(con, y, project)
        ctx["pl"]["losses"] = [0, 0]
        by = pl_compute(ctx, y)
        if y < int(as_of[:4]):
            profit += by[58]["cum"] - max(by[58]["cum"], 0) * mz_settings()["rate"]
        else:
            m = int(as_of[5:7])
            pre = sum(by[58]["m"][:m])
            profit += pre - max(pre, 0) * mz_settings()["rate"]
    return dict(cash=cash, deb=deb + op.get("Debitor", 0), kred=kred + op.get("Kreditor", 0), stock=stock,
                fa_gross=gross, fa_acc=acc, fa=gross - acc, oth_a=op.get("Digər aktiv", 0), oth_l=op.get("Digər öhdəlik", 0),
                cap=eq["Nizamnamə kapitalı"] + eq["Əlavə kapital"], div=eq["Dividend"],
                loan=eq["Təsisçi borcu (daxil olan)"] + eq["Təsisçi borcu (qaytarılan)"],
                profit=profit * 100 + op.get("Yığılmış mənfəət", 0), as_of=as_of)


def mz_bs(con, as_of, project=None):
    b = mz_bs_data(con, as_of or dt.date.today().isoformat(), project)
    A = [("Pul vəsaitləri (bank)", b["cash"]), ("Debitor borcları", b["deb"]), ("Ehtiyat və mallar", b["stock"]),
         ("Əsas vəsaitlər (qalıq dəyər)", b["fa"]), ("Digər aktivlər", b["oth_a"])]
    L = [("Kreditor borcları", b["kred"]), ("Təsisçi borcu", b["loan"]), ("Digər öhdəliklər", b["oth_l"])]
    E = [("Nizamnamə və əlavə kapital", b["cap"]), ("Yığılmış mənfəət (P&L + açılış)", b["profit"]),
         ("Dividend", b["div"])]
    ta, tl, te = sum(v for _, v in A), sum(v for _, v in L), sum(v for _, v in E)
    rows = [{"c": ["AKTİVLƏR", ""], "k": "grp"}] + [{"c": ["   " + n, _az(v)], "k": ""} for n, v in A]
    rows += [{"c": ["Cəmi aktivlər", _az(ta)], "k": "tot"}, {"c": ["ÖHDƏLİKLƏR", ""], "k": "grp"}]
    rows += [{"c": ["   " + n, _az(v)], "k": ""} for n, v in L]
    rows += [{"c": ["Cəmi öhdəliklər", _az(tl)], "k": "tot"}, {"c": ["KAPİTAL", ""], "k": "grp"}]
    rows += [{"c": ["   " + n, _az(v)], "k": ""} for n, v in E]
    rows += [{"c": ["Cəmi kapital", _az(te)], "k": "tot"},
             {"c": ["Cəmi öhdəlik + kapital", _az(tl + te)], "k": "tot"},
             {"c": ["Balanslaşmayan fərq (aktiv − passiv)", _az(ta - tl - te)], "k": "risk" if abs(ta - tl - te) > 100 else ""}]
    return dict(cols=[["Maddə", 0], [b["as_of"], 1]], rows=rows,
                note="İkiqat yazılış yoxdur: fərq sıfırdan çoxdursa 'Kapital / ƏV > Açılış qalıqları' bölməsində başlanğıc "
                     "qalıqları (pul, debitor, kreditor, yığılmış mənfəət) daxil edin.")


# ---------------------------------------------------------------- Ratio
def mz_ratio(con, d1, d2, project=None):
    d2 = d2 or dt.date.today().isoformat()
    d1 = d1 or f"{d2[:4]}-01-01"
    days = (dt.date.fromisoformat(d2) - dt.date.fromisoformat(d1)).days + 1
    t = mz_pl_period(con, d1, d2, project)
    b = mz_bs_data(con, d2, project)
    P = project or ""
    rev_v = con.execute("SELECT COALESCE(SUM(total),0) FROM inv WHERE dir='S' AND date BETWEEN ? AND ? AND (?='' OR project=?)", (d1, d2, P, P)).fetchone()[0] / 100
    pur_v = con.execute("SELECT COALESCE(SUM(total),0) FROM inv WHERE dir='G' AND date BETWEEN ? AND ? AND (?='' OR project=?)", (d1, d2, P, P)).fetchone()[0] / 100
    rev, cogs, gross = t[6], t[14], t[20]
    opex, amort, fin, ebt, net = t[22], t[46], t[55], t[58], t[58] - max(t[58], 0) * mz_settings()["rate"]
    ebit = ebt + fin
    ebitda = ebit + amort
    cash, deb, stock, kred = b["cash"] / 100, b["deb"] / 100, b["stock"] / 100, b["kred"] / 100
    cl = kred + (b["oth_l"] / 100)
    ta = (b["cash"] + b["deb"] + b["stock"] + b["fa"] + b["oth_a"]) / 100
    te = (b["cap"] + b["profit"] + b["div"]) / 100
    tl = cl + b["loan"] / 100

    def r(a, bb, pct=False, d=2):
        if not bb:
            return None
        v = a / bb
        return round(v * 100, d) if pct else round(v, d)

    R = [("LİKVİDLİK", None, ""),
         ("Cari likvidlik", r(cash + deb + stock, cl), "(pul + debitor + ehtiyat) / qısamüddətli öhdəlik"),
         ("Sürətli likvidlik", r(cash + deb, cl), "(pul + debitor) / qısamüddətli öhdəlik"),
         ("Mütləq likvidlik", r(cash, cl), "pul / qısamüddətli öhdəlik"),
         ("FƏALİYYƏT (günlə)", None, ""),
         ("Debitor ödəmə müddəti (DSO)", r(deb * days, rev_v), "debitor × gün / gedən qaimə cəmi (ƏDV daxil)"),
         ("Kreditor ödəmə müddəti (DPO)", r(kred * days, pur_v), "kreditor × gün / gələn qaimə cəmi (ƏDV daxil)"),
         ("Ehtiyatın dövriyyə müddəti (DIO)", r(stock * days, cogs), "ehtiyat × gün / maya dəyəri"),
         ("Pul dövriyyə dövrü (CCC)", None if None in (r(deb * days, rev_v), r(kred * days, pur_v), r(stock * days, cogs)) else
          round(r(deb * days, rev_v) + r(stock * days, cogs) - r(kred * days, pur_v), 2), "DSO + DIO − DPO"),
         ("Aktivlərin dövriyyəsi (dəfə)", r(rev, ta), "satış / cəmi aktivlər"),
         ("MƏNFƏƏTLİLİK (%)", None, ""),
         ("Brüt marja", r(gross, rev, True), "brüt qazanc / satış"),
         ("EBITDA marjası", r(ebitda, rev, True), "(vergidən öncə + maliyyə xərci + amortizasiya) / satış"),
         ("Əməliyyat marjası (EBIT)", r(ebit, rev, True), "(vergidən öncə + maliyyə xərci) / satış"),
         ("Xalis marja", r(net, rev, True), "vergidən sonra / satış"),
         ("ROA", r(net, ta, True), "xalis mənfəət / cəmi aktivlər (dövr üzrə)"),
         ("ROE", r(net, te, True), "xalis mənfəət / kapital (dövr üzrə)"),
         ("KAPİTAL QURULUŞU", None, ""),
         ("Öhdəlik / kapital", r(tl, te), "cəmi öhdəlik / kapital"),
         ("Öhdəlik / aktiv (%)", r(tl, ta, True), "cəmi öhdəlik / cəmi aktivlər")]
    rows = [{"c": [n, v if v is not None else "", f], "k": "grp" if v is None and not f else ""} for n, v, f in R]
    return dict(cols=[["Göstərici", 0], ["Dəyər", 1], ["Düstur", 0]], rows=rows,
                note=f"{d1} – {d2} ({days} gün)" + (f" | layihə: {project}" if project else "") +
                     f" | satış {rev:,.2f}, maya {cogs:,.2f}, xalis {net:,.2f}")


# ---------------------------------------------------------------- debitor-kreditor ilə sinxron
def mz_sync_txn(con):
    """Qaimə və bank hərəkətlərini mövcud Debitor/Kreditor cədvəlinə (txn) ötürür (mənbə 'MAZA')."""
    with con:
        con.execute("DELETE FROM txn WHERE src='MAZA'")
        rows, n = [], 0
        for dr, d, voen, doc, base, vat in con.execute(
                "SELECT dir,date,voen,doc_no,SUM(base),SUM(vat) FROM inv GROUP BY dir,date,voen,doc_no"):
            neg = base < 0 or (base == 0 and vat < 0)
            kind = ("SQ" if neg else "S") if dr == "S" else ("AQ" if neg else "A")
            b, v = abs(base), abs(vat)
            for amt, part in ((b, "E"), (v, "V")):
                if amt:
                    n += 1
                    rows.append(("MAZA", n, kind, d, doc, voen, amt, "", 18, part))
        for d, voen, ref, kind, inn, out in con.execute(
                "SELECT date,voen,ref,kind,inflow,outflow FROM bank WHERE voen<>'' AND kind IN ('Debitor ödəmə','Kreditor ödəmə')"):
            n += 1
            if kind == "Debitor ödəmə":
                rows.append(("MAZA", n, "D" if inn else "DQ", d, ref, voen, inn or out, "", 18, ""))
            else:
                rows.append(("MAZA", n, "C" if out else "CQ", d, ref, voen, out or inn, "", 18, ""))
        con.executemany("INSERT INTO txn(src,row_no,kind,date,doc_no,voen,amount,note,vat,part) VALUES (?,?,?,?,?,?,?,?,?,?)", rows)
    return len(rows)


# ---------------------------------------------------------------- generik cədvəl redaktəsi
MZ_FIELDS = {
    "projects": [("name", "Layihə adı", "text"), ("note", "Qeyd", "text")],
    "map_class": [("mtype", "Qayda növü", "text"), ("pattern", "Açar (VÖEN / söz / kateqoriya)", "text"),
                  ("side", "Tərəf (G gələn, S gedən, N qeyri-rezident, *)", "text"), ("cls", "Sinif", "text"),
                  ("pl_line", "P&L maddəsi", "text"), ("project", "Layihə", "text"), ("prio", "Prioritet (az = əvvəl)", "num")],
    "map_bank": [("mtype", "Qayda növü", "text"), ("pattern", "Açar", "text"), ("kind", "Hərəkət növü", "text"),
                 ("cf_section", "Pul axını bölməsi", "text"), ("pl_line", "P&L maddəsi", "text"),
                 ("project", "Layihə", "text"), ("prio", "Prioritet", "num")],
    "map_item": [("raw", "Qaimədəki mal adı", "text"), ("std", "Vahid ad", "text"), ("unit", "Vahid", "text")],
    "nonres": [("kind", "Növ (idxal/ixrac/xidmet)", "text"), ("company", "Şirkət", "text"), ("doc_no", "İnvoys №", "text"),
               ("doc_date", "İnvoys tarixi", "date"), ("purpose", "Təyinat", "text"), ("category", "Kateqoriya", "text"),
               ("project", "Layihə", "text"), ("ccy", "Valyuta", "text"), ("amount_ccy", "Məbləğ (valyuta)", "num"),
               ("amount", "Məbləğ (AZN)", "num"), ("vat", "ƏDV (AZN)", "num"), ("cls", "Sinif", "text"),
               ("pl_line", "P&L maddəsi", "text"), ("item", "Mal", "text"), ("hs", "Mal kodu", "text"),
               ("gb", "GB / sənəd", "text"), ("note", "Qeyd", "text")],
    "assets": [("name", "Əsas vəsait", "text"), ("category", "Kateqoriya", "text"), ("cost", "İlkin dəyər", "num"),
               ("acc_amort", "Yığılmış amortizasiya (başlanğıc)", "num"), ("rate", "Amortizasiya faizi (7 və ya 0.07)", "num"),
               ("start_date", "Mədaxil / qalıq tarixi", "date"), ("project", "Layihə", "text")],
    "equity": [("date", "Tarix", "date"), ("kind", "Növ", "text"), ("amount", "Məbləğ (AZN)", "num"), ("note", "Qeyd", "text")],
    "opening": [("acct", "Hesab", "text"), ("amount", "Açılış qalığı (AZN)", "num")],
}
MZ_LABELS = {"projects": "Layihələr", "map_class": "Qaimə təsnifatı", "map_bank": "Bank qaydaları", "map_item": "Mal adları",
             "nonres": "Qeyri-rezident", "assets": "Əsas vəsaitlər", "equity": "Kapital hərəkətləri", "opening": "Açılış qalıqları"}
MZ_MONEY = {"nonres": ["amount", "vat"], "assets": ["cost", "acc_amort"], "equity": ["amount"], "opening": ["amount"]}
MZ_TEXT_INTS = {"prio"}


def mz_options(con):
    prj = [r[0] for r in con.execute("SELECT name FROM projects ORDER BY name")]
    return {"mtype": ["voen", "name", "item", "subcat"], "cls": MZ_CLS, "side": ["*", "G", "S", "N"], "pl_line": mz_pl_labels(),
            "project": prj, "kind": MZ_KINDS, "cf_section": MZ_CF, "category": MZ_ASSET_CATS,
            "acct": MZ_OPEN_ACCTS}


def mz_fields_meta(con, tb):
    o = mz_options(con)
    out = []
    for k, l, t in MZ_FIELDS[tb]:
        opts = o.get(k)
        if tb == "map_bank" and k == "mtype":
            opts = ["cat", "text", "voen"]
        if tb == "equity" and k == "kind":
            opts = MZ_EQ_KINDS
        if tb == "nonres" and k == "kind":
            opts = ["idxal", "ixrac", "xidmet"]
        if tb == "assets" and k == "category":
            opts = MZ_ASSET_CATS
        out.append({"k": k, "l": l, "t": t, "o": opts})
    return out


def mz_list(con, tb):
    if tb not in MZ_FIELDS:
        raise ValueError("naməlum cədvəl")
    cols = ["id"] + [f[0] for f in MZ_FIELDS[tb]]
    order = {"nonres": "doc_date DESC,id DESC", "equity": "date,id", "map_class": "prio,id", "map_bank": "prio,id"}.get(tb, "id")
    rows = []
    for r in con.execute(f"SELECT {','.join(cols)} FROM {tb} ORDER BY {order} LIMIT 3000"):
        d = dict(zip(cols, r))
        for c in MZ_MONEY.get(tb, []):
            d[c] = _az(d[c])
        rows.append(d)
    return {"fields": mz_fields_meta(con, tb), "rows": rows, "table": tb}


def mz_save(con, tb, row):
    if tb not in MZ_FIELDS:
        raise ValueError("naməlum cədvəl")
    keys = [f[0] for f in MZ_FIELDS[tb]]
    vals = {}
    for k, l, t in MZ_FIELDS[tb]:
        v = row.get(k)
        if k in MZ_MONEY.get(tb, []):
            v = _q(v)
        elif t == "num":
            v = _fl(v)
            v = int(v) if k in MZ_TEXT_INTS else v
        else:
            v = _s(v)
        vals[k] = v
    if tb == "nonres":
        vals["manual"] = 1
        keys = keys + ["manual"]
    rid = row.get("id")
    with con:
        if rid:
            con.execute(f"UPDATE {tb} SET " + ",".join(f"{k}=?" for k in keys) + " WHERE id=?", [vals[k] for k in keys] + [rid])
        else:
            con.execute(f"INSERT OR REPLACE INTO {tb}({','.join(keys)}) VALUES({','.join('?' * len(keys))})", [vals[k] for k in keys])
        if tb == "opening":
            pass
    if tb in ("map_class", "map_bank", "map_item", "projects", "nonres"):
        mz_classify(con)
    return True


def mz_delete(con, tb, rid):
    if tb not in MZ_FIELDS:
        raise ValueError("naməlum cədvəl")
    with con:
        con.execute(f"DELETE FROM {tb} WHERE id=?", (rid,))
    if tb in ("map_class", "map_bank"):
        mz_classify(con)
    return True


def mz_unmapped(con):
    inv = []
    for dr, voen, name, n, tot, sample, sub in con.execute(
            "SELECT dir,voen,MAX(name),COUNT(*),SUM(total),MAX(COALESCE(NULLIF(item,''),subcat,cat)),"
            "MAX(COALESCE(NULLIF(subcat,''),cat)) FROM inv WHERE COALESCE(cls,'')='' GROUP BY dir,voen ORDER BY SUM(ABS(total)) DESC LIMIT 300"):
        inv.append({"dir": dr, "voen": voen, "name": name or "", "n": n, "total": _az(tot), "sample": sample or "", "sub": sub or ""})
    bk = []
    for voen, n, inn, out, sample in con.execute(
            "SELECT voen,COUNT(*),SUM(inflow),SUM(outflow),MAX(details) FROM bank WHERE kind='Təsnifatsız' "
            "GROUP BY CASE WHEN voen<>'' THEN voen ELSE substr(details,1,40) END ORDER BY SUM(inflow+outflow) DESC LIMIT 300"):
        bk.append({"voen": voen or "", "n": n, "in": _az(inn), "out": _az(out), "sample": (sample or "")[:140]})
    nr = [{"id": r[0], "company": r[1], "category": r[2], "amount": _az(r[3])} for r in con.execute(
        "SELECT id,company,category,amount FROM nonres WHERE COALESCE(cls,'')='' LIMIT 200")]
    return {"inv": inv, "bank": bk, "nonres": nr}


def mz_counts(con):
    return {t: con.execute(f"SELECT COUNT(*) FROM {t}").fetchone()[0] for t in
            ("inv", "bank", "nonres", "payroll", "assets", "map_class", "map_bank", "projects", "equity")} | {
        "inv_unclassified": con.execute("SELECT COUNT(*) FROM inv WHERE COALESCE(cls,'')=''").fetchone()[0],
        "bank_unclassified": con.execute("SELECT COUNT(*) FROM bank WHERE kind='Təsnifatsız'").fetchone()[0]}


def mz_export(path, con, year, d1, d2, project=None):
    d2 = d2 or dt.date.today().isoformat()
    d1 = d1 or f"{d2[:4]}-01-01"
    wb = Workbook()
    wb.remove(wb.active)
    for title, rep in (("Mənfəət-zərər", mz_pl(con, year, project)), ("Pul axını", mz_cf(con, year, project)),
                       ("Balans", mz_bs(con, d2, project)), ("Ratio", mz_ratio(con, d1, d2, project)),
                       ("VÖEN üzrə", mz_voen(con, d1, d2, project)), ("Ehtiyat", mz_stock(con, d1, d2, project))):
        ws = wb.create_sheet(title)
        ws.append([c[0] for c in rep["cols"]])
        _style_header(ws, len(rep["cols"]))
        for r in rep["rows"]:
            ws.append(r["c"])
            for c in range(1, len(rep["cols"]) + 1):
                cell = ws.cell(row=ws.max_row, column=c)
                cell.font = Font(name=FONT, bold=r["k"] in ("grp", "tot"))
                cell.border = BORDER
                if rep["cols"][c - 1][1] and isinstance(cell.value, (int, float)):
                    cell.number_format = MONEY_FMT
                if r["k"] in ("grp", "tot"):
                    cell.fill = TOTAL_FILL
        ws.column_dimensions["A"].width = 46
        for c in range(2, len(rep["cols"]) + 1):
            ws.column_dimensions[get_column_letter(c)].width = 17
        ws.freeze_panes = "B2"
    ws = wb.create_sheet("Qeyri-rezident")
    ws.append(["Növ", "Şirkət", "İnvoys №", "İnvoys tarixi", "Təyinat", "Kateqoriya", "Layihə", "AZN", "Sinif", "P&L maddəsi"])
    _style_header(ws, 10)
    for r in con.execute("SELECT kind,company,doc_no,doc_date,purpose,category,project,amount,cls,pl_line FROM nonres ORDER BY doc_date"):
        ws.append(list(r[:7]) + [_az(r[7]), r[8], r[9]])
    wb.save(path)


# ---------------------------------------------------------------- HTTP marşrutları
def maza_get(db, path, q):
    con = mz_connect(db)
    try:
        ep = path[len("/api/maza/"):]
        P = q.get("project") or None
        d1 = parse_date(q.get("from")) if q.get("from") else None
        d2 = parse_date(q.get("to")) if q.get("to") else None
        today = dt.date.today().isoformat()
        year = int(q.get("year") or (d2 or today)[:4])
        if ep == "init":
            ys = sorted({int(r[0]) for r in con.execute(
                "SELECT DISTINCT substr(date,1,4) FROM inv UNION SELECT DISTINCT substr(date,1,4) FROM bank "
                "UNION SELECT DISTINCT substr(doc_date,1,4) FROM nonres") if r[0] and r[0].isdigit()} | {dt.date.today().year})
            return {"projects": [r[0] for r in con.execute("SELECT name FROM projects ORDER BY name")],
                    "labels": MZ_LABELS, "counts": mz_counts(con), "years": ys, "folder": str(BASE / "maza_data"),
                    "settings": mz_settings(), "cls": MZ_CLS_LABEL}
        if ep == "list":
            return mz_list(con, q["table"])
        if ep == "unmapped":
            return mz_unmapped(con)
        if ep == "voen":
            return mz_voen(con, d1, d2, P)
        if ep == "stock":
            r = mz_stock(con, d1, d2, P)
            r.pop("total", None)
            return r
        if ep == "pl":
            r = mz_pl(con, year, P)
            r.pop("by", None)
            return r
        if ep == "cf":
            return mz_cf(con, year, P)
        if ep == "bs":
            return mz_bs(con, d2 or today, P)
        if ep == "ratio":
            return mz_ratio(con, d1, d2, P)
        if ep == "items":
            return mz_items_unmapped(con)
        if ep == "checks":
            return mz_checks(con)
        if ep == "export":
            tmp = tempfile.mktemp(suffix=".xlsx")
            mz_export(tmp, con, year, d1, d2, P)
            data = open(tmp, "rb").read()
            os.remove(tmp)
            return ("bin", data, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", f"maza_hesabat_{year}.xlsx")
        raise ValueError("naməlum yol")
    finally:
        con.close()


def maza_post(db, path, body, bump=lambda: None):
    con = mz_connect(db)
    try:
        ep = path[len("/api/maza/"):]
        if ep == "import":
            rep = mz_import_path(con, body.get("path") or str(BASE / "maza_data"))
            bump()
            return {"ok": True, "report": [list(r) for r in rep], "counts": mz_counts(con)}
        if ep == "classify":
            mz_seed(con)
            mz_classify(con)
            bump()
            return {"ok": True, "counts": mz_counts(con)}
        if ep == "templates":
            return {"ok": True, "files": [str(f) for f in make_maza_templates(BASE / "sablonlar" / "maza")]}
        if ep == "sync":
            n = mz_sync_txn(con)
            bump()
            return {"ok": True, "rows": n}
        if ep == "save":
            mz_save(con, body["table"], body.get("row") or {})
            bump()
            return {"ok": True}
        if ep == "delete":
            mz_delete(con, body["table"], body["id"])
            bump()
            return {"ok": True}
        if ep == "settings":
            cfg = load_cfg()
            s = mz_settings(cfg)
            for k in ("rate", "loss1", "loss2", "cap1", "cap2"):
                if k in body:
                    s[k] = _fl(body[k])
            if s["rate"] > 1:
                s["rate"] /= 100.0
            if "own_voen" in body:
                s["own_voen"] = _s(body["own_voen"])
            cfg["maza"] = s
            save_cfg(cfg)
            return {"ok": True}
        raise ValueError("naməlum yol")
    finally:
        con.close()


# ---------------------------------------------------------------- əlavələr: mal uyğunlaşdırma, yoxlama, şablonlar, ehtiyat nüsxə
def mz_items_unmapped(con):
    """E-qaimədəki mal adları (alış: ehtiyat/mal sinfi; satış: hamısı) ki, hələ 'Mal adları' cədvəlində yoxdur."""
    mapped = {pl_fold(r[0]) for r in con.execute("SELECT raw FROM map_item")}
    mapped |= {pl_fold(r[0]) for r in con.execute("SELECT std FROM map_item") if r[0]}
    out = {}
    for dr, item, unit, base in con.execute(
            "SELECT dir,item,unit,base FROM inv WHERE item<>'' AND (cls IN ('ehtiyat','mal') OR dir='S')"):
        k = pl_fold(item)
        if k in mapped:
            continue
        e = out.setdefault(k, dict(item=item, unit=unit or "", g=0, s=0, total=0))
        e["g" if dr == "G" else "s"] += 1
        e["total"] += base
    rows = [dict(item=e["item"], unit=e["unit"], g=e["g"], s=e["s"], total=_az(e["total"]),
                 role="alış + satış" if e["g"] and e["s"] else "alış" if e["g"] else "satış")
            for e in sorted(out.values(), key=lambda e: -abs(e["total"]))]
    return {"items": rows[:500], "more": max(0, len(rows) - 500)}


def mz_checks(con):
    c = mz_counts(con)
    today = dt.date.today().isoformat()
    out = []

    def add(label, n, hint=""):
        out.append({"label": label, "n": n, "ok": not n, "hint": hint})

    add("Təsnifatsız qaimə sətirləri", c["inv_unclassified"], "Mapping > Təsnifatsızlar")
    add("Təsnifatsız bank hərəkətləri", c["bank_unclassified"], "Mapping > Təsnifatsızlar")
    add("Təsnifatsız qeyri-rezident sənədlər",
        con.execute("SELECT COUNT(*) FROM nonres WHERE COALESCE(cls,'')=''").fetchone()[0], "Qeyri-rezident cədvəli")
    add("Uyğunlaşdırılmamış mal adları", len(mz_items_unmapped(con)["items"]), "Mapping > Təsnifatsızlar")
    add("Mənfi ehtiyat qalığı olan mallar",
        sum(1 for r in mz_stock(con, None, today)["rows"] if r["k"] == "risk"), "Ehtiyat tabı")
    b = mz_bs_data(con, today)
    diff = (b["cash"] + b["deb"] + b["stock"] + b["fa"] + b["oth_a"]) - (
        b["kred"] + b["loan"] + b["oth_l"] + b["cap"] + b["profit"] + b["div"])
    add("Balans fərqi (AZN)", round(abs(diff) / 100, 2) if abs(diff) > 100 else 0, "Kapital / ƏV > Açılış qalıqları")
    return {"rows": out}


MZ_TEMPLATES = {
    "gelen": ("sablon_gelen_qaime.xlsx", "Gələn qaimələr (təchizatçılardan). Hər sətir = qaimənin bir mal/xidmət sətri."),
    "geden": ("sablon_geden_qaime.xlsx", "Gedən qaimələr (müştərilərə). Hər sətir = bir qaimə."),
    "bank": ("sablon_bank_cixaris.xlsx", "Bank çıxarışı. VÖEN və 'Ümumi kateqoriya' doldurulubsa avtomatik növə bağlanır "
                                         "(debitor, kreditor, bank xərci, əmək haqqı, vergi, kredit, təsisçi ...)."),
    "ecustom": ("sablon_idxal_ecustom.xlsx", "Gömrük idxal bəyannamələri (e-customs)."),
    "gemi": ("sablon_qeyri_rezident_xidmet.xlsx", "Qeyri-rezident xidmət invoysları (gəmi icarəsi və s.)."),
    "payroll": ("sablon_emek_haqqi.xlsx", "Aylıq əmək haqqı cədvəli (işəgötürən payları ilə)."),
}


def make_maza_templates(folder):
    """M.A.Z.A idxal şablonları: başlıqlar oxuyucunun tanıdığı adlarla eynidir."""
    folder = Path(folder)
    folder.mkdir(parents=True, exist_ok=True)
    out = []
    for kind, (fname, note) in MZ_TEMPLATES.items():
        spec = MZ_SPECS[kind]
        heads = [alts[0] for alts in spec["map"].values()]
        wb = Workbook()
        ws = wb.active
        ws.title = "Data"
        ws.append(heads)
        _style_header(ws, len(heads))
        for i in range(1, len(heads) + 1):
            ws.column_dimensions[get_column_letter(i)].width = max(14, min(38, len(heads[i - 1]) + 4))
        ws.freeze_panes = "A2"
        info = wb.create_sheet("Təlimat")
        info.column_dimensions["A"].width = 110
        req = [spec["map"][k][0] for k in spec["req"]]
        for ln in (note, "", "Faylı maza_data qovluğuna atın, sonra M.A.Z.A > İdxal > 'Oxu' basın.",
                   "Başlıq sətrini silməyin; sütunların sırası sərbəstdir, lazımsız sütunları boş qoya bilərsiniz.",
                   "Məcburi sütunlar: " + ", ".join(req),
                   "'Layihə' sütunu doldurulubsa sətir həmin layihəyə aid olur (Terminal layihəsi və s.).",
                   "Təkrar oxunduqda eyni fayl əvvəlcə silinib yenidən yazılır; başqa faylda olan eyni sətir dublikat sayılıb keçilir."):
            info.append([ln])
        wb.save(folder / fname)
        out.append(folder / fname)
    return out


def backup_db(db_path, keep=20):
    """SQLite bazasının tutarlı ehtiyat nüsxəsi (backups/ qovluğu, son 'keep' nüsxə saxlanır)."""
    db_path = Path(db_path)
    folder = BASE / "backups"
    folder.mkdir(parents=True, exist_ok=True)
    dst = folder / f"{db_path.stem}_{dt.datetime.now():%Y%m%d_%H%M%S}.db"
    src = sqlite3.connect(str(db_path))
    out = sqlite3.connect(str(dst))
    try:
        src.backup(out)
    finally:
        out.close()
        src.close()
    for p in sorted(folder.glob(f"{db_path.stem}_*.db"))[:-keep]:
        try:
            p.unlink()
        except OSError:
            pass
    return dst


# ----------------------------------------------------------------------------
# Əmr sətri
# ----------------------------------------------------------------------------

def main():
    ap = argparse.ArgumentParser(description="Debitor-Kreditor")
    ap.add_argument("--sablon", action="store_true", help="Excel şablonlarını yarat")
    ap.add_argument("--yenile", action="store_true", help="papkaları oxu")
    ap.add_argument("--export", metavar="FAYL.xlsx", help="debitor/kreditor cədvəllərini Excelə ixrac et")
    ap.add_argument("--voen-yoxla", action="store_true", help="VÖEN-ləri portaldan yoxla (ad, ƏDV, riskli)")
    ap.add_argument("--hamisi", action="store_true", help="--voen-yoxla ilə: artıq yoxlanmışları da yenilə")
    ap.add_argument("--uzanti", nargs="?", const="uzanti", metavar="QOVLUQ", help="Chrome uzantısını (Vergi alətləri) qovluğa çıxar")
    ap.add_argument("--tk", action="store_true", help="köhnə Tkinter pəncərəsi (standart: brauzer interfeysi)")
    ap.add_argument("--server", action="store_true", help="yalnız uzantıdan qəbul (pəncərəsiz)")
    ap.add_argument("--tarix-dan", metavar="GG.AA.İİİİ", help="aralığın başlanğıcı (əvvəlki qalıq ayrıca göstərilir)")
    ap.add_argument("--tarix", metavar="GG.AA.İİİİ", help="qalıq tarixi (boşdursa hamısı)")
    ap.add_argument("--backup", action="store_true", help="bazanın ehtiyat nüsxəsini backups/ qovluğuna yaz")
    ap.add_argument("--maza-sablon", action="store_true", help="M.A.Z.A idxal şablonlarını yarat")
    ap.add_argument("--maza-import", metavar="YOL", help="M.A.Z.A: fayl və ya qovluğu oxu (xlsx/xlsm/xlsb)")
    ap.add_argument("--maza-export", metavar="FAYL.xlsx", help="M.A.Z.A hesabatlarını (P&L, pul axını, balans, ratio ...) Excelə ixrac et")
    ap.add_argument("--il", type=int, help="--maza-export üçün il")
    ap.add_argument("--layihe", help="--maza-export üçün layihə adı")
    a = ap.parse_args()

    if a.uzanti:
        print("Uzantı çıxarıldı:", extract_extension(a.uzanti), "\n(Chrome -> chrome://extensions -> Developer mode -> Load unpacked)")
        return
    if a.backup:
        print("Ehtiyat nüsxə:", backup_db(load_cfg()["db"]))
        return
    if a.maza_sablon:
        for f in make_maza_templates(BASE / "sablonlar" / "maza"):
            print("Yaradıldı:", f)
        return
    if a.maza_import:
        con = mz_connect(load_cfg()["db"])
        for r in mz_import_path(con, a.maza_import, print):
            print(f"  {r[0]} | {r[1]} | yeni: {r[2]}, dublikat: {r[3]}")
        print(mz_counts(con))
        return
    if a.maza_export:
        con = mz_connect(load_cfg()["db"])
        d2 = parse_date(a.tarix) if a.tarix else None
        d1 = parse_date(a.tarix_dan) if a.tarix_dan else None
        mz_export(a.maza_export, con, a.il or int((d2 or dt.date.today().isoformat())[:4]), d1, d2, a.layihe)
        print("İxrac edildi:", a.maza_export)
        return
    if a.server:
        start_server(load_cfg()["db"], lambda n: print(f"Uzantıdan {n} sətir alındı"))
        print(f"Dinlənir: 127.0.0.1:{LISTEN_PORT} (dayandırmaq üçün Ctrl+C)")
        try:
            while True:
                time.sleep(3600)
        except KeyboardInterrupt:
            return
    if not (a.sablon or a.yenile or a.export or a.voen_yoxla):
        if a.tk:
            run_gui()
            return
        if serve(sys.modules[__name__]) is None:
            return
        try:
            while True:
                time.sleep(3600)
        except KeyboardInterrupt:
            return
    cfg = load_cfg()
    if a.sablon:
        for f in make_templates(BASE / "sablonlar"):
            print("Yaradıldı:", f)
    if a.yenile or a.export or a.voen_yoxla:
        con = connect(cfg["db"])
        if a.yenile or a.export or a.voen_yoxla:
            s = scan(con, cfg, print)
            print(f"Yeni: {s['yeni']}, yenilənən: {s['yenilenen']}, dəyişməyən: {s['kecilen']}, "
                  f"silinən: {s['silinen']}, sətir: {s['setir']}, xəta: {len(s['xetalar'])}, "
                  f"dublikat qrup: {s['dublikat']}, müddət: {s['muddet']:.2f} san")
            for f, r, m in s["xetalar"][:50]:
                print(f"  XƏTA {f}, sətir {r}: {m}")
        if a.voen_yoxla:
            todo = voens_to_check(con, a.hamisi)
            if todo:
                check_voens(cfg["db"], cfg.get("sessionId", ""), todo,
                            lambda d, t, e, m: print(f"VÖEN {d}/{t}" + (f" xəta {e}: {m}" if e else ""), end="\r"))
                print()
        if a.export:
            as_of = parse_date(a.tarix) if a.tarix else None
            if a.tarix and not as_of:
                sys.exit("Tarix formatı səhvdir (GG.AA.İİİİ).")
            start = parse_date(a.tarix_dan) if a.tarix_dan else None
            export_xlsx(a.export, con, as_of, start)
            print("İxrac edildi:", a.export)


if __name__ == "__main__":
    main()

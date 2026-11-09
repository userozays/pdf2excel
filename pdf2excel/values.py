"""Hücre metnini sayı/tarihe çevirme ve Excel biçimini geri metne render etme.

Kural: bir metin yalnızca Excel'e yazılacak değer + number_format, Excel'in
göstereceği metni kaynakla BİREBİR üretiyorsa dönüştürülür. Aksi halde metin kalır.
"""
from __future__ import annotations

import re
from datetime import date, datetime
from decimal import Decimal, InvalidOperation

SEPS = {"tr": (".", ","), "en": (",", ".")}  # (binlik, ondalık)

CUR_PREFIX = r"(?P<pre>[₺$€£]\s?)?"
CUR_SUFFIX = r"(?P<suf>\s?(?:TL|TRY|USD|EUR|GBP|₺|\$|€|£))?"


def _num_re(thou: str, dec: str) -> re.Pattern:
    t, d = re.escape(thou), re.escape(dec)
    body = rf"(?P<int>\d{{1,3}}(?:{t}\d{{3}})+|\d+)(?:{d}(?P<frac>\d+))?"
    return re.compile(
        rf"^(?P<paren>\()?(?P<sign>[-+−])?(?P<pct1>%)?{CUR_PREFIX}{body}(?P<pct2>%)?{CUR_SUFFIX}(?P<trail>-)?(?(paren)\))$"
    )


NUM_RE = {loc: _num_re(*SEPS[loc]) for loc in SEPS}

DATE_RE = re.compile(
    r"^(?:(?P<d>\d{1,2})(?P<s1>[./-])(?P<m>\d{1,2})(?P=s1)(?P<y>\d{4})"
    r"|(?P<Y>\d{4})-(?P<M>\d{2})-(?P<D>\d{2}))"
    r"(?:\s(?P<H>\d{2}):(?P<Mi>\d{2})(?::(?P<S>\d{2}))?)?$"
)

# Belge yerel ayarı için kanıt: yalnız tek anlamlı biçimler sayılır
TR_EVIDENCE = re.compile(r"^[-(]?\d{1,3}(?:\.\d{3})*,\d+\)?-?$")
EN_EVIDENCE = re.compile(r"^[-(]?\d{1,3}(?:,\d{3})*\.\d+\)?-?$")


def detect_locale(texts) -> str | None:
    tr = en = 0
    for t in texts:
        t = t.strip().replace("₺", "").replace("TL", "").strip()
        # "1.234" ve "1,234" iki anlamlı; kanıt sayılmaz
        if re.fullmatch(r"\d{1,3}[.,]\d{3}", t):
            continue
        if TR_EVIDENCE.match(t):
            tr += 1
        elif EN_EVIDENCE.match(t):
            en += 1
    if tr == en:
        return None
    return "tr" if tr > en else "en"


def _q(s: str) -> str:
    return '"' + s.replace('"', '""') + '"' if s else ""


def parse(text: str, locale: str | None):
    """(değer, number_format) döndürür ya da None (metin olarak kalsın)."""
    t = text.strip()
    if not t or t != text or len(t) > 40:
        return None
    if locale is None:
        # yerel ayar bilinmiyorsa yalnız ayraçsız tam sayılar ve tarihler
        if re.fullmatch(r"-?\d+", t):
            locale = "en"
        else:
            return _parse_date(t)
    m = NUM_RE[locale].match(t)
    if not m:
        return _parse_date(t)
    thou, dec = SEPS[locale]
    g = m.groupdict()
    if g["pct1"] and g["pct2"]:
        return None
    if len(g["int"]) > 1 and g["int"][0] == "0":  # 00123: hesap/sicil no, sıfırlar korunmalı
        return None
    digits = g["int"].replace(thou, "") + (g["frac"] or "")
    if len(digits.lstrip("0")) > 15:  # Excel 15 anlamlı basamak
        return None
    try:
        val = Decimal(g["int"].replace(thou, "") + ("." + g["frac"] if g["frac"] else ""))
    except InvalidOperation:
        return None
    neg_styles = [k for k in ("paren", "trail") if g[k]] + (["sign"] if g["sign"] in ("-", "−") else [])
    if len(neg_styles) > 1 or g["sign"] == "+":
        return None
    pct = g["pct1"] or g["pct2"]
    if pct:
        val = val / 100
    core = ("#,##0" if thou in g["int"] else "0") + ("." + "0" * len(g["frac"]) if g["frac"] else "")
    pre = (g["pre"] or "")
    suf = (g["suf"] or "")
    pos = (_q(pre)) + ("%" if g["pct1"] else "") + core + ("%" if g["pct2"] else "") + _q(suf)
    fmt = pos
    if neg_styles:
        val = -val
        if g["paren"]:
            fmt = f"{pos};({pos})"
        elif g["trail"]:
            fmt = f"{pos};{pos}-"
        elif g["sign"] == "−":
            fmt = f'{pos};"−"{pos}'
        elif val == 0:  # "-0,00" -> Excel "-" göstermez
            return None
    num = int(val) if val == val.to_integral_value() and not g["frac"] and not pct else float(val)
    if isinstance(num, float) and Decimal(repr(num)) != val:
        return None
    return num, fmt


def _parse_date(t: str):
    m = DATE_RE.match(t)
    if not m:
        return None
    g = m.groupdict()
    try:
        if g["y"]:
            d0 = (int(g["y"]), int(g["m"]), int(g["d"]))
            s = g["s1"]
            fmt = ("dd" if len(g["d"]) == 2 else "d") + s + ("mm" if len(g["m"]) == 2 else "m") + s + "yyyy"
        else:
            d0 = (int(g["Y"]), int(g["M"]), int(g["D"]))
            fmt = "yyyy-mm-dd"
        if g["H"]:
            v = datetime(*d0, int(g["H"]), int(g["Mi"]), int(g["S"] or 0))
            fmt += " hh:mm" + (":ss" if g["S"] else "")
        else:
            v = datetime(*d0)
    except ValueError:
        return None
    if v.year < 1900:
        return None
    return v, fmt


# ---------------------------------------------------------------- render
def _split_sections(fmt: str):
    out, cur, q = [], "", False
    for ch in fmt:
        if ch == '"':
            q = not q
        if ch == ";" and not q:
            out.append(cur)
            cur = ""
        else:
            cur += ch
    out.append(cur)
    return out


def _tokens(sec: str):
    """Bölümü (literal, kod) parçalarına ayırır."""
    parts = re.findall(r'"((?:[^"]|"")*)"|([^"]+)', sec)
    return [(a.replace('""', '"'), None) if a or (not b) else (None, b) for a, b in parts]


def render(value, fmt: str, locale: str) -> str:
    """openpyxl'den okunan değer + biçimi, Excel'in (verilen yerel ayarda) göstereceği metne çevirir.
    Yalnız bu uygulamanın ürettiği biçim alt kümesini destekler."""
    if isinstance(value, str):
        return value
    if isinstance(value, (datetime, date)):
        if not isinstance(value, datetime):
            value = datetime(value.year, value.month, value.day)
        out = fmt
        for k, v in (("yyyy", f"{value.year:04d}"), ("hh", f"{value.hour:02d}"), ("ss", f"{value.second:02d}")):
            out = out.replace(k, v)
        # dakika ':' sonrasındaki mm; gün/ay
        out = re.sub(r"(?<=:)mm", f"{value.minute:02d}", out)
        out = out.replace("dd", f"{value.day:02d}").replace("mm", f"{value.month:02d}")
        out = re.sub(r"(?<![0-9])d(?![0-9])", str(value.day), out, count=1)
        out = re.sub(r"(?<![0-9])m(?![0-9])", str(value.month), out, count=1)
        return out
    thou, dec = SEPS[locale]
    secs = _split_sections(fmt or "General")
    if (fmt or "General") == "General":
        d = Decimal(repr(value)) if isinstance(value, float) else Decimal(value)
        s = format(d.normalize(), "f") if isinstance(value, float) else str(value)
        return s.replace(".", dec)
    neg = value < 0
    sec = secs[1] if neg and len(secs) > 1 else secs[0]
    v = abs(value)
    # tek bölümlü biçimde Excel eksi işaretini en başa koyar
    out = ["-"] if neg and len(secs) == 1 else []
    for lit, code in _tokens(sec):
        if lit is not None:
            out.append(lit)
            continue
        m = re.search(r"(#,##0|0)(\.(0+))?", code)
        pct = "%" in code
        nd = len(m.group(3)) if m and m.group(3) else 0
        x = Decimal(repr(v)) if isinstance(v, float) else Decimal(v)
        if pct:
            x *= 100
        s = f"{x:,.{nd}f}" if m and m.group(1) == "#,##0" else f"{x:.{nd}f}"
        s = s.replace(",", "\0").replace(".", dec).replace("\0", thou)
        if m:
            code = code[: m.start()] + s + code[m.end():]
        out.append(code)
    return "".join(out)

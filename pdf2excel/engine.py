"""PDF -> tablo modeli.

Akış: sayfa -> (çizgili tablolar + serbest satırlar) -> sütun ızgarası -> satır modeli.
Her hücre PDF'teki kelimelerin kendisinden kurulur; hiçbir karakter üretilmez veya atılmaz.
"""
from __future__ import annotations

import statistics
from collections import Counter
from dataclasses import dataclass, field

import pdfplumber

from . import values

LIGATURES = {"ﬀ": "ff", "ﬁ": "fi", "ﬂ": "fl", "ﬃ": "ffi", "ﬄ": "ffl", "ﬅ": "st", "ﬆ": "st"}
# Eski Windows-1254 fontlarında sık görülen yanlış eşleme
TR_FIX = {"ý": "ı", "Ý": "İ", "þ": "ş", "Þ": "Ş", "ð": "ğ", "Ð": "Ğ"}


@dataclass
class Options:
    mode: str = "tek"            # "tek": tüm sayfalar tek sayfada, "sayfa": sayfa başına sheet
    locale: str = "auto"         # "auto" | "tr" | "en"
    convert: bool = True         # sayı/tarih dönüştür
    join_wrapped: bool = True    # alt satıra kaymış açıklamaları üst satıra ekle
    tr_fix: str = "auto"         # "auto" | "on" | "off"
    skip_repeated: bool = True   # her sayfada tekrarlanan üst/alt bilgiyi yalnız ilk sayfada tut


@dataclass
class Seg:
    x0: float
    x1: float
    top: float
    bottom: float
    size: float
    text: str


@dataclass
class Line:
    page: int
    top: float
    bottom: float
    segs: list[Seg]


@dataclass
class Row:
    page: int
    cells: dict[int, str]
    notes: dict[int, set] = field(default_factory=dict)
    kind: str = "grid"          # grid | ruled
    top: float = 0.0
    bottom: float = 0.0
    size: float = 10.0
    merges: list = field(default_factory=list)  # (c0, rspan, cspan) ruled tablo için, satır göreli
    aligned: bool = True        # her parça kendi sütununun içinde başlıyor mu
    moved: list = field(default_factory=list)   # [(sayfa, Counter)] başka sayfadan gelen karakterler


@dataclass
class Sheet:
    name: str
    rows: list[Row]
    merges: list = field(default_factory=list)  # (r0, c0, r1, c1) 0 tabanlı
    typed: dict = field(default_factory=dict)   # (r, c) -> (value, fmt)


@dataclass
class Result:
    sheets: list[Sheet]
    page_chars: dict[int, Counter]
    locale: str | None
    info: dict
    warnings: list[str]


def _norm(s: str, trmap: dict | None) -> str:
    for a, b in LIGATURES.items():
        s = s.replace(a, b)
    if trmap:
        s = s.translate(str.maketrans(trmap))
    return s


# ------------------------------------------------------------ satır/segment
def _lines_from_words(words, page_no: int) -> list[Line]:
    words = sorted(words, key=lambda w: (round(w["top"], 1), w["x0"]))
    lines: list[list] = []
    for w in words:
        mid = (w["top"] + w["bottom"]) / 2
        for ln in reversed(lines[-3:]):
            top = min(x["top"] for x in ln)
            bot = max(x["bottom"] for x in ln)
            if top <= mid <= bot:
                ln.append(w)
                break
        else:
            lines.append([w])
    out = []
    for ln in lines:
        ln.sort(key=lambda w: w["x0"])
        segs: list[Seg] = []
        for w in ln:
            size = w.get("size") or (w["bottom"] - w["top"])
            if segs and w["x0"] - segs[-1].x1 <= 0.5 * max(size, segs[-1].size):
                s = segs[-1]
                s.x1 = max(s.x1, w["x1"])
                s.top, s.bottom = min(s.top, w["top"]), max(s.bottom, w["bottom"])
                s.text += " " + w["text"]
            else:
                segs.append(Seg(w["x0"], w["x1"], w["top"], w["bottom"], size, w["text"]))
        out.append(Line(page_no, min(s.top for s in segs), max(s.bottom for s in segs), segs))
    out.sort(key=lambda l: l.top)
    return out


# ------------------------------------------------------------ sütun ızgarası
def _looks_typed(text: str) -> bool:
    return values.parse(text, "tr") is not None or values.parse(text, "en") is not None


def _columns(lines: list[Line]) -> list[tuple[float, float]]:
    multi = [l for l in lines if len(l.segs) >= 2]
    if not multi:
        return [(0.0, 1e9)]
    # Sütunları gerçek tablo satırlarından (sayı/tarih içeren çok parçalı satırlar) çıkar;
    # müşteri/adres gibi üst bilgi satırları sütun aralarını kapatmasın.
    base = [l for l in multi if any(_looks_typed(s.text) for s in l.segs)]
    if len(base) < 3:
        base = multi
    sizes = [s.size for l in base for s in l.segs]
    em = statistics.median(sizes)
    R = 2  # yarım punto çözünürlük
    width = int(max(s.x1 for l in lines for s in l.segs) * R) + 4
    cov = [0] * width
    for l in base:
        for s in l.segs:
            for i in range(int(s.x0 * R), int(s.x1 * R) + 1):
                cov[i] += 1
    tol = int(0.04 * len(base))
    runs, start = [], None
    for i, c in enumerate(cov + [0]):
        if c > tol and start is None:
            start = i
        elif c <= tol and start is not None:
            runs.append([start / R, i / R])
            start = None
    # dar boşlukları sütun arası saymayız
    merged = []
    for r in runs:
        if merged and r[0] - merged[-1][1] < 0.4 * em:
            merged[-1][1] = r[1]
        else:
            merged.append(r)
    # boşluk bölgesine tamamen düşen segmentler seyrek bir sütuna işarettir
    extra = []
    for l in multi:
        for s in l.segs:
            if not any(s.x0 < c[1] and s.x1 > c[0] for c in merged):
                extra.append([s.x0, s.x1])
    extra.sort()
    for e in extra:
        if merged and any(e[0] < c[1] and e[1] > c[0] for c in merged):
            for c in merged:
                if e[0] < c[1] and e[1] > c[0]:
                    c[0], c[1] = min(c[0], e[0]), max(c[1], e[1])
                    break
        else:
            merged.append(e)
    merged.sort()
    return [tuple(c) for c in merged]


def _assign(seg: Seg, cols) -> tuple[int, bool]:
    best, bov, spans = 0, -1.0, 0
    for i, (a, b) in enumerate(cols):
        ov = min(seg.x1, b) - max(seg.x0, a)
        if ov > 0:
            spans += 1
        if ov > bov:
            best, bov = i, ov
    if bov <= 0:
        mid = (seg.x0 + seg.x1) / 2
        best = min(range(len(cols)), key=lambda i: min(abs(mid - cols[i][0]), abs(mid - cols[i][1])))
    elif spans > 1:
        # birden çok sütuna taşan (çoğunlukla sola yaslı) metin başladığı sütuna aittir
        start = [i for i, (a, b) in enumerate(cols) if a <= seg.x0 < b]
        if start:
            best = start[0]
    return best, spans > 1


def _grid_rows(lines: list[Line], cols) -> list[Row]:
    rows = []
    for l in lines:
        cells: dict[int, str] = {}
        notes: dict[int, set] = {}
        aligned = True
        for s in l.segs:
            c, spans = _assign(s, cols)
            aligned &= cols[c][0] - 1 <= s.x0 <= cols[c][1]
            if c in cells:
                cells[c] += " " + s.text
                notes.setdefault(c, set()).add("birlesik")
            else:
                cells[c] = s.text
            if spans and len(l.segs) > 1:
                notes.setdefault(c, set()).add("tasma")
        rows.append(Row(l.page, cells, notes, "grid", l.top, l.bottom,
                        statistics.median(s.size for s in l.segs), aligned=aligned))
    return rows


def _join_wrapped(rows: list[Row], locale) -> int:
    """Çok satırlı kayıtları tek satırda toplar.

    Çapa: sayı/tarih içeren satır (ör. tarih + tutar + bakiye). Anahtar sütunu (çapaların
    çoğunda dolu en soldaki sütun) boş olan, yalnız metin içeren satırlar, aralarında boşluk
    olmadan zincirlenerek ulaşılabilen en yakın çapaya eklenir. Açıklama çapanın altında,
    üstünde ya da ortasında olabilir. Eklenen satır sayısını döndürür."""
    def typed(r):
        return any(values.parse(t, locale) is not None for t in r.cells.values())

    def mid(r):
        return (r.top + r.bottom) / 2

    anchors = [r for r in rows if r.kind == "grid" and typed(r)]
    if len(anchors) < 3:
        return 0
    fill = Counter(c for r in anchors for c in r.cells)
    keys = [c for c, n in fill.items() if n >= 0.5 * len(anchors)]
    if not keys:
        return 0
    key = min(keys)
    is_anchor = {id(r) for r in anchors}

    def attachable(r):
        return r.kind == "grid" and id(r) not in is_anchor and key not in r.cells and r.aligned

    def fits(r, a):  # metin, çapanın sayı/tarih hücresinin üstüne yazılmasın
        return all(c not in a.cells or values.parse(a.cells[c], locale) is None for c in r.cells)

    target: dict[int, Row] = {}
    n = len(rows)
    for i, r in enumerate(rows):
        if not attachable(r):
            continue
        found = []
        for step in (-1, 1):
            j, prev = i, r
            while 0 <= j + step < n:
                j += step
                q = rows[j]
                if q.page != r.page or abs(mid(q) - mid(prev)) > 1.9 * max(q.size, prev.size):
                    break
                if id(q) in is_anchor:
                    found.append(q)
                    break
                if not attachable(q):
                    break
                prev = q
        found = [a for a in found if fits(r, a)]
        if found:
            target[id(r)] = min(found, key=lambda a: abs(mid(a) - mid(r)))

    # Sayfa sonunda bölünen açıklama: sayfanın başındaki, kendi sayfasında çapaya ulaşamayan
    # devam satırları önceki sayfanın son kaydına eklenir ve kontrol için işaretlenir.
    cross = set()
    for i, r in enumerate(rows):
        if id(r) in target or not attachable(r):
            continue
        before = [q for q in rows[:i] if q.page == r.page]
        if any(not (id(q) in target or id(q) in cross) for q in before):
            continue  # sayfanın en başındaki devam satırları değil
        prev = [q for q in rows[:i] if q.page == r.page - 1 and id(q) in is_anchor and key in q.cells]
        if prev and fits(r, prev[-1]):
            target[id(r)] = prev[-1]
            cross.add(id(r))

    groups: dict[int, list[Row]] = {}
    for r in rows:
        if id(r) in target:
            groups.setdefault(id(target[id(r)]), []).append(r)
    out = []
    for r in rows:
        if id(r) in target:
            continue
        extra = groups.get(id(r))
        if extra:
            parts = sorted([r] + extra, key=lambda x: (x.page, x.top))
            cells: dict[int, str] = {}
            notes: dict[int, set] = {}
            for p in parts:
                for c, t in p.cells.items():
                    cells[c] = cells[c] + " " + t if c in cells else t
                    notes.setdefault(c, set()).update(p.notes.get(c, set()))
                    if p is not r:
                        notes[c].add("sayfa" if id(p) in cross else "devam")
                if p.page != r.page:
                    r.moved.append((p.page, Counter(ch for t in p.cells.values() for ch in t if not ch.isspace())))
            r.cells = dict(sorted(cells.items()))
            r.notes = {c: v for c, v in notes.items() if v}
            same = [p for p in parts if p.page == r.page]
            r.top, r.bottom = same[0].top, same[-1].bottom
        out.append(r)
    rows[:] = out
    return len(target)


# ------------------------------------------------------------ çizgili tablolar
def _snap(vals, tol=2.0):
    vals = sorted(vals)
    out = []
    for v in vals:
        if not out or v - out[-1] > tol:
            out.append(v)
    return out


def _idx(edges, v, tol=2.0):
    return min(range(len(edges)), key=lambda i: abs(edges[i] - v))


def _ruled_rows(table, words, page_no) -> tuple[list[Row], list]:
    xs = _snap([c[0] for c in table.cells] + [c[2] for c in table.cells])
    ys = _snap([c[1] for c in table.cells] + [c[3] for c in table.cells])
    nrows = len(ys) - 1
    rows = [Row(page_no, {}, {}, "ruled", ys[i], ys[i + 1]) for i in range(nrows)]
    used = set()
    for (x0, t, x1, b) in table.cells:
        c0, c1 = _idx(xs, x0), _idx(xs, x1)
        r0, r1 = _idx(ys, t), _idx(ys, b)
        if c1 <= c0 or r1 <= r0:
            continue
        inside = [i for i, w in enumerate(words) if i not in used
                  and x0 <= (w["x0"] + w["x1"]) / 2 <= x1 and t <= (w["top"] + w["bottom"]) / 2 <= b]
        used.update(inside)
        if inside:
            lines = _lines_from_words([words[i] for i in inside], page_no)
            rows[r0].cells[c0] = " ".join(" ".join(s.text for s in l.segs) for l in lines)
            rows[r0].size = statistics.median(w.get("size", 10) for w in (words[i] for i in inside))
        if c1 - c0 > 1 or r1 - r0 > 1:
            rows[r0].merges.append((c0, r1 - r0, c1 - c0))
    leftover = [w for i, w in enumerate(words) if i not in used]
    return rows, leftover


# ------------------------------------------------------------ ana giriş
def convert(path: str, opts: Options | None = None) -> Result:
    opts = opts or Options()
    warnings: list[str] = []
    info = Counter()
    page_chars: dict[int, Counter] = {}
    per_page: list[tuple[int, list[Line], list[Row]]] = []

    with pdfplumber.open(path) as pdf:
        raw_text = "".join(c["text"] for p in pdf.pages for c in p.chars)
        trmap = None
        if opts.tr_fix == "on" or (
            opts.tr_fix == "auto" and any(ch in raw_text for ch in "ýþðÝÞÐ")
            and not any(ch in raw_text for ch in "ışğİŞĞ")
        ):
            trmap = TR_FIX
            info["tr_duzeltme"] = sum(raw_text.count(k) for k in TR_FIX)

        for pno, page in enumerate(pdf.pages, start=1):
            p = page.dedupe_chars()
            info["mukerrer_karakter"] += len(page.chars) - len(p.chars)
            page_chars[pno] = Counter(
                ch for c in p.chars for ch in _norm(c["text"], trmap) if not ch.isspace()
            )
            words = p.extract_words(x_tolerance_ratio=0.15, y_tolerance=3, extra_attrs=["size"])
            for w in words:
                w["text"] = _norm(w["text"], trmap)
            if not words:
                warnings.append(f"Sayfa {pno}: metin katmanı yok (taranmış olabilir, OCR gerekir).")
                per_page.append((pno, [], []))
                continue
            ruled_rows: list[Row] = []
            tables = [t for t in p.find_tables({"vertical_strategy": "lines", "horizontal_strategy": "lines"})
                      if len(t.rows) >= 2 and max(len(r.cells) for r in t.rows) >= 2]
            free = words
            for t in tables:
                # gölgelendirme kutularından doğan, çoğu boş "tablolar"ı atla
                filled = sum(1 for (x0, tp, x1, b) in t.cells if any(
                    x0 <= (w["x0"] + w["x1"]) / 2 <= x1 and tp <= (w["top"] + w["bottom"]) / 2 <= b for w in free))
                if filled < 0.5 * len(t.cells):
                    continue
                rows, free = _ruled_rows(t, free, pno)
                ruled_rows.extend(rows)
                info["cizgili_tablo"] += 1
            per_page.append((pno, _lines_from_words(free, pno), ruled_rows))

    locale = None if opts.locale == "auto" else opts.locale

    # Her sayfada aynı yerde tekrar eden üst/alt bilgi satırları sütun tespitini bozmasın
    def sig(l):
        return (round(l.top), " ".join(s.text for s in l.segs))
    seen = Counter(s for _, ls, _ in per_page for s in {sig(l) for l in ls})
    repeated = {s for s, k in seen.items() if k >= 2 and k >= 0.5 * len(per_page)}
    if opts.skip_repeated and repeated:
        # ilk görüldüğü sayfada kalır; sonraki sayfalardaki kopyalar atlanır ve
        # doğrulamada o sayfanın beklenen karakterlerinden düşülür
        first: dict = {}
        for i, (pno, ls, rr) in enumerate(per_page):
            keep = []
            for l in ls:
                s = sig(l)
                if s in repeated and first.setdefault(s, pno) != pno:
                    page_chars[pno] -= Counter(ch for seg in l.segs for ch in seg.text if not ch.isspace())
                    info["atlanan_tekrar_satir"] += 1
                else:
                    keep.append(l)
            per_page[i] = (pno, keep, rr)

    def build(pages) -> list[Row]:
        lines = [l for _, ls, _ in pages for l in ls]
        cols = _columns([l for l in lines if sig(l) not in repeated] or lines)
        rows: list[Row] = []
        for pno, ls, rr in pages:
            page_rows = _grid_rows(ls, cols) + rr
            page_rows.sort(key=lambda r: r.top)
            rows.extend(page_rows)
        return rows

    groups = [per_page] if opts.mode == "tek" else [[pp] for pp in per_page]
    names = ["Tablo"] if opts.mode == "tek" else [f"Sayfa {pp[0]}" for pp in per_page]
    sheets = [Sheet(n, build(g)) for n, g in zip(names, groups)]

    if locale is None and opts.locale == "auto":
        locale = values.detect_locale(t for s in sheets for r in s.rows for t in r.cells.values())
        if locale is None:
            warnings.append("Sayı biçimi (TR/EN) belirlenemedi; ayraçlı sayılar metin olarak bırakıldı.")
    elif opts.locale in ("tr", "en"):
        seen_loc = values.detect_locale(t for s in sheets for r in s.rows for t in r.cells.values())
        if seen_loc and seen_loc != opts.locale:
            names = {"tr": "1.234,56", "en": "1,234.56"}
            warnings.append(
                f"Seçilen sayı biçimi ({names[opts.locale]}) bu belgeyle uyuşmuyor; belgede {names[seen_loc]} "
                "biçimi kullanılmış. Yanlış değer üretmemek için bu sayılar metin olarak bırakıldı. "
                "\"Otomatik\" seçerseniz sayıya çevrilirler.")

    for s in sheets:
        if opts.join_wrapped:
            info["birlestirilen_satir"] += _join_wrapped(s.rows, locale)
        _finalize(s, locale, opts, info)

    return Result(sheets, page_chars, locale, dict(info), warnings)


def _finalize(sheet: Sheet, locale, opts: Options, info: Counter):
    # boş sütunları at, merge'leri mutlak konuma çevir, tipleri belirle
    used = sorted({c for r in sheet.rows for c in r.cells} |
                  {c + k for r in sheet.rows for (c, _, cs) in r.merges for k in range(cs)})
    remap = {c: i for i, c in enumerate(used)}
    for ri, r in enumerate(sheet.rows):
        r.cells = {remap[c]: t for c, t in r.cells.items()}
        r.notes = {remap[c]: n for c, n in r.notes.items() if c in remap}
        for c0, rs, cs in r.merges:
            a, b = remap[c0], remap[c0 + cs - 1]
            sheet.merges.append((ri, a, ri + rs - 1, b))
        if opts.convert:
            for c, t in r.cells.items():
                pv = values.parse(t, locale)
                if pv is not None:
                    sheet.typed[(ri, c)] = pv
                    info["tarih" if hasattr(pv[0], "year") else "sayi"] += 1

"""Anlamsal kontroller. Bakiye zinciri: önceki bakiye + işlem tutarı = yeni bakiye.

Karakter kontrolü içeriğin eksiksiz olduğunu kanıtlar; bakiye zinciri ise satırların
eksiksiz, doğru sırada ve doğru sütunlarda olduğunu kanıtlar (ekstrelerde)."""
from __future__ import annotations

from decimal import Decimal

from openpyxl.utils import get_column_letter


def _norm(t: str) -> str:
    return t.replace("İ", "i").replace("I", "ı").lower()


NAMES = {
    "bakiye": ("bakiye", "balance"),
    "tutar": ("tutar", "amount"),
    "borc": ("borç", "borc", "debit"),
    "alacak": ("alacak", "credit"),
}
TOTAL_WORDS = ("toplam", "total", "genel")


def _find_header(sheet):
    for ri, row in enumerate(sheet.rows):
        found = {}
        for c, t in row.cells.items():
            if (ri, c) in sheet.typed or len(t) > 30:
                continue
            n = _norm(t)
            for k, words in NAMES.items():
                if k not in found and any(w in n for w in words):
                    found[k] = c
        if "bakiye" in found and ("tutar" in found or ("borc" in found and "alacak" in found)):
            return ri, found
    return None, None


def _num(sheet, ri, c):
    v = sheet.typed.get((ri, c))
    if v is None or isinstance(v[0], (str,)) or hasattr(v[0], "year") or "%" in v[1]:
        return None
    return Decimal(repr(v[0])) if isinstance(v[0], float) else Decimal(v[0])


def _run(items, sign_debit):
    """items: [(ri, amount_parts, balance)] sırayla. Hata listesi ve kontrol sayısı döner."""
    running, pending, errors, checked, opening = None, Decimal(0), [], 0, None
    for ri, (tutar, borc, alacak), bal in items:
        amt = None
        if tutar is not None:
            amt = tutar
        elif borc is not None or alacak is not None:
            amt = (alacak or 0) + sign_debit * (borc or 0)
        if amt is not None:
            pending += amt
        if bal is not None:
            if running is None:
                opening = bal - pending if amt is not None else bal
            else:
                checked += 1
                if running + pending != bal:
                    errors.append((ri, running + pending, bal))
            running, pending = bal, Decimal(0)
    return errors, checked, opening, running


def _fmt(d, loc):
    if d is None:
        return None
    s = f"{d:,.2f}"
    return s.replace(",", "\0").replace(".", ",").replace("\0", ".") if loc == "tr" else s


def balance_chain(sheet, loc: str | None = "tr") -> dict | None:
    hr, cols = _find_header(sheet)
    if hr is None:
        return None
    items = []
    for ri in range(hr + 1, len(sheet.rows)):
        row = sheet.rows[ri]
        if any(any(w in _norm(t) for w in TOTAL_WORDS) for c, t in row.cells.items() if (ri, c) not in sheet.typed):
            continue
        parts = tuple(_num(sheet, ri, cols[k]) if k in cols else None for k in ("tutar", "borc", "alacak"))
        bal = _num(sheet, ri, cols["bakiye"])
        if bal is None and all(p is None for p in parts):
            continue
        items.append((ri, parts, bal))
    if sum(1 for *_, b in items if b is not None) < 2:
        return None
    # Yön (eskiden yeniye / yeniden eskiye) ve borcun işareti (pozitif yazılıp düşülen ya da
    # zaten negatif yazılan) belgeden belgeye değişir: en az hata veren yorum seçilir.
    # Yeniden eskiye sıralı ekstrede liste ters çevrilince aynı kural geçerli olur.
    trials = []
    for order, seq in (("eskiden yeniye", items), ("yeniden eskiye", items[::-1])):
        for sign in ((1,) if "tutar" in cols else (-1, 1)):
            errs, checked, opening, closing = _run(seq, sign)
            trials.append((len(errs), -checked, order, sign, errs, checked, opening, closing))
    trials.sort(key=lambda t: (t[0], t[1]))
    nerr, _, order, sign, errs, checked, opening, closing = trials[0]
    return {
        "sheet": sheet.name,
        "columns": {k: get_column_letter(c + 1) for k, c in cols.items()},
        "order": order,
        "checked": checked,
        "opening": _fmt(opening, loc),
        "closing": _fmt(closing, loc),
        "errors": [{"row": ri, "ref": f"{get_column_letter(cols['bakiye'] + 1)}{ri + 1}",
                    "expected": _fmt(exp, loc), "actual": _fmt(act, loc)} for ri, exp, act in errs],
    }

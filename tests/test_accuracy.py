"""Üretilen PDF'leri çevirir ve Excel'i bilinen doğru tabloyla hücre hücre karşılaştırır.
Çalıştır:  .venv\\Scripts\\python.exe -m tests.test_accuracy"""
from __future__ import annotations

import sys
from datetime import datetime
from decimal import Decimal
from pathlib import Path

from openpyxl import load_workbook

from pdf2excel import Options, convert_file
from pdf2excel.values import render
from tests import samples

OUT = Path(__file__).parent / "out"
OUT.mkdir(exist_ok=True)
fails = 0


def check(cond, msg):
    global fails
    if not cond:
        fails += 1
        print("  HATA:", msg)


def shown(ws, r, c, loc):
    cell = ws.cell(r, c)
    return render(cell.value, cell.number_format, loc) if cell.value is not None else ""


def find_header(ws, head):
    for row in ws.iter_rows():
        vals = [c.value for c in row]
        if all(h in vals for h in head):
            return row[0].row, {h: vals.index(h) + 1 for h in head}
    return None, None


def num(s, loc):
    s = s.replace("TL", "").strip()
    neg = s.startswith("(")
    s = s.strip("()")
    pct = s.endswith("%")
    s = s.rstrip("%")
    s = s.replace(".", "").replace(",", ".") if loc == "tr" else s.replace(",", "")
    v = Decimal(s) * (-1 if neg else 1)
    return v / 100 if pct else v


def run(name, make, opts=None):
    pdf, xlsx = OUT / f"{name}.pdf", OUT / f"{name}.xlsx"
    truth = make(pdf)
    res, rep = convert_file(str(pdf), str(xlsx), opts)
    print(f"[{name}] doğrulama={'OK' if rep['ok'] else 'SORUN'} hücre={rep['cells']} "
          f"yerel={rep['locale']} bilgi={rep['info']} işaretli={len(rep['flags'])}")
    for p in rep["pages"]:
        check(p["ok"], f"sayfa {p['page']} karakter farkı: {p['diff']}")
    for e in rep["value_errors"][:5]:
        check(False, f"değer {e}")
    for w in rep["warnings"]:
        print("  uyarı:", w)
    return truth, load_workbook(xlsx), rep


def test_ekstre():
    (head, truth, totals), wb, rep = run("ekstre_tr", samples.ekstre)
    ws = wb["Tablo"]
    hr, cm = find_header(ws, head)
    check(hr is not None, "başlık satırı bulunamadı")
    if not hr:
        return
    got = []
    for r in range(hr + 1, ws.max_row + 1):
        v = ws.cell(r, cm["Tarih"]).value
        if isinstance(v, datetime):
            got.append(r)
    check(len(got) == len(truth), f"işlem sayısı {len(got)} != {len(truth)}")
    bad = 0
    for r, t in zip(got, truth):
        for h in head:
            s = shown(ws, r, cm[h], "tr")
            if s != t[h]:
                bad += 1
                if bad <= 5:
                    check(False, f"satır {r} {h}: {s!r} != {t[h]!r}")
        for h in ("Borç", "Alacak", "Bakiye"):
            v = ws.cell(r, cm[h]).value
            if t[h]:
                check(isinstance(v, (int, float)) and Decimal(repr(v)) == num(t[h], "tr"),
                      f"satır {r} {h} sayı değil/yanlış: {v!r}")
    check(bad == 0, f"{bad} hücre metni farklı")
    tot = [r for r in range(1, ws.max_row + 1) if ws.cell(r, cm["Açıklama"]).value == "TOPLAM"]
    check(len(tot) == 1 and [shown(ws, tot[0], cm[h], "tr") for h in ("Borç", "Alacak")] == totals["TOPLAM"],
          "TOPLAM satırı yanlış")
    print(f"  {len(got)} işlem satırı x {len(head)} sütun karşılaştırıldı")


def test_fatura():
    (head, items), wb, rep = run("fatura_cizgili", samples.fatura)
    ws = wb["Tablo"]
    hr, cm = find_header(ws, head)
    check(hr is not None, "başlık satırı bulunamadı")
    if not hr:
        return
    for i, item in enumerate(items):
        for h, t in zip(head, item):
            check(shown(ws, hr + 1 + i, cm[h], "tr") == t, f"{h}: {shown(ws, hr + 1 + i, cm[h], 'tr')!r} != {t!r}")
    tutar = ws.cell(hr + 1, cm["Tutar"]).value
    check(isinstance(tutar, (int, float)) and abs(tutar - 64900) < 1e-9, f"Tutar sayı değil: {tutar!r}")
    merges = {str(m) for m in ws.merged_cells.ranges}
    print("  birleşik hücreler:", sorted(merges))
    check(len(merges) == 3, "3 birleşik hücre bekleniyordu (Kalemler, Fiyatlandırma, Genel toplam)")


def test_en():
    (head, data), wb, rep = run("rapor_en", samples.rapor_en)
    ws = wb["Tablo"]
    hr, cm = find_header(ws, head)
    check(hr is not None, "başlık satırı bulunamadı")
    if not hr:
        return
    for i, row in enumerate(data):
        for h, t in zip(head, row):
            check(shown(ws, hr + 1 + i, cm[h], "en") == t, f"{h}: {shown(ws, hr + 1 + i, cm[h], 'en')!r} != {t!r}")
            if h != "Region":
                v = ws.cell(hr + 1 + i, cm[h]).value
                check(isinstance(v, (int, float)) and Decimal(repr(v)) == num(t, "en"), f"{h} sayı yanlış: {v!r}")


def test_zor():
    _, wb, rep = run("zor", samples.zor)
    ws = wb["Tablo"]
    got = {ws.cell(r, 1).value: ws.cell(r, 2) for r in range(1, ws.max_row + 1)}
    check(ws["A1"].value == "HESAP DÖKÜMÜ", f"sahte kalın tekrar etti: {ws['A1'].value!r}")
    check(rep["info"].get("mukerrer_karakter", 0) >= 12, "mükerrer karakterler atılmadı")
    check("İşlem yeri" in got and got["İşlem yeri"].value == "İstanbul Şişli Dağıtım", "Türkçe onarım çalışmadı")
    expect_text = {"Hesap No": "00123456", "Kart No": "4543 1234 5678 9012",
                   "Kart No2": "4543123456789012", "Formül gibi": "=SUM(A1:A3)", "Sıfır": "-0,00"}
    for k, v in expect_text.items():
        cell = got.get(k)
        check(cell is not None and cell.value == v and cell.data_type == "s", f"{k} metin kalmalıydı: {cell and cell.value!r}")
    expect_num = {"Oran": 0.18, "Eksi sonda": -1250, "Parantez": -2000.5, "Tutar TL": 12345.67}
    for k, v in expect_num.items():
        cell = got.get(k)
        check(cell is not None and isinstance(cell.value, (int, float)) and abs(cell.value - v) < 1e-9,
              f"{k} sayı olmalıydı: {cell and cell.value!r}")
    dt = got.get("Tarih saat")
    check(dt is not None and dt.value == datetime(2026, 11, 7, 14, 35), f"tarih-saat: {dt and dt.value!r}")


def test_bakiye_zinciri():
    _, _, rep = run("ekstre_tr", samples.ekstre)
    b = rep["balance"]
    check(len(b) == 1 and b[0]["checked"] == 69 and not b[0]["errors"],
          f"doğru ekstrede zincir tutmalıydı: {b and (b[0]['checked'], b[0]['errors'][:2])}")
    _, _, rep = run("ekstre_ters", lambda p: samples.ekstre(p, seed=3, ters=True))
    b = rep["balance"]
    check(len(b) == 1 and not b[0]["errors"] and b[0]["order"] == "yeniden eskiye",
          f"ters sıralı ekstre tanınmadı: {b and (b[0]['order'], len(b[0]['errors']))}")
    _, wb, rep = run("ekstre_bozuk", lambda p: samples.ekstre(p, seed=5, bozuk=30))
    b = rep["balance"]
    errs = b[0]["errors"] if b else []
    check(not rep["ok"] and errs, "bozuk bakiye yakalanmadı")
    if errs:
        ws = wb["Tablo"]
        hr, cm = find_header(ws, ["Tarih", "Bakiye"])
        data_rows = [r for r in range(hr + 1, ws.max_row + 1) if isinstance(ws.cell(r, cm["Tarih"]).value, datetime)]
        check(errs[0]["ref"] == f"E{data_rows[30]}", f"kopukluk yanlış satırda: {errs[0]['ref']} != E{data_rows[30]}")
        print(f"  bozuk ekstre: {len(errs)} kopukluk, ilki {errs[0]['ref']} (beklenen {errs[0]['expected']}, PDF {errs[0]['actual']})")


def test_sayfa_modu():
    _, wb, rep = run("ekstre_sayfa", lambda p: samples.ekstre(p, n=120, seed=7), Options(mode="sayfa"))
    check(len(wb.sheetnames) >= 3, f"sayfa başına sheet bekleniyordu: {wb.sheetnames}")
    check(rep["ok"], "sayfa modunda doğrulama başarısız")
    check(rep["balance"] and not any(b["errors"] for b in rep["balance"]),
          f"sayfa modunda bakiye zinciri: {[(b['sheet'], len(b['errors'])) for b in rep['balance']]}")


if __name__ == "__main__":
    for t in (test_ekstre, test_fatura, test_en, test_zor, test_bakiye_zinciri, test_sayfa_modu):
        t()
    print("SONUÇ:", "TÜM TESTLER GEÇTİ" if fails == 0 else f"{fails} HATA")
    sys.exit(1 if fails else 0)

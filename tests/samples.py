"""Bilinen veriden test PDF'leri üretir. Her örnek, PDF'e basılan doğru tabloyu (ground truth) döndürür."""
from __future__ import annotations

import random
from datetime import date, timedelta
from decimal import Decimal
from pathlib import Path

from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas

pdfmetrics.registerFont(TTFont("Arial", r"C:\Windows\Fonts\arial.ttf"))
pdfmetrics.registerFont(TTFont("ArialB", r"C:\Windows\Fonts\arialbd.ttf"))
W, H = A4


def tr(d: Decimal) -> str:
    s = f"{d:,.2f}"
    return s.replace(",", "X").replace(".", ",").replace("X", ".")


def en(d: Decimal) -> str:
    return f"{d:,.2f}"


def _wrap(text, font, size, width):
    lines, cur = [], ""
    for w in text.split(" "):
        t = (cur + " " + w).strip()
        if pdfmetrics.stringWidth(t, font, size) > width and cur:
            lines.append(cur)
            cur = w
        else:
            cur = t
    return lines + [cur]


ACIKLAMALAR = [
    "ATM PARA ÇEKME", "EFT GELEN", "KREDİ KARTI ÖDEMESİ", "MAAŞ ÖDEMESİ ŞİRKET A.Ş.",
    "FAST GİDEN AYŞE YILMAZ KİRA ÖDEMESİ KASIM AYI DAİRE 4",
    "POS ALIŞVERİŞ MİGROS TİC. A.Ş. İSTANBUL ÜSKÜDAR ŞUBESİ",
    "HAVALE GELEN", "FATURA ÖDEMESİ İSKİ", "ÇEK TAHSİLATI 0012345",
    "OTOMATİK ÖDEME TÜRK TELEKOM İNTERNET HİZMET BEDELİ",
]


def ekstre(path: Path, n=70, seed=1, ters=False, bozuk=None):
    """Çizgisiz, 3+ sayfalık banka ekstresi. Seyrek Borç/Alacak, alta kayan açıklamalar.
    ters: yeniden eskiye sıralı basılır. bozuk: bu sıradaki satırın bakiyesi 1,00 fazla basılır."""
    rnd = random.Random(seed)
    c = canvas.Canvas(str(path), pagesize=A4)
    cols = {"tarih": 40, "aciklama": 105, "borc": 395, "alacak": 470, "bakiye": 555}
    desc_w = 225  # açıklama en geniş Borç tutarına değmeden biter
    head = ["Tarih", "Açıklama", "Borç", "Alacak", "Bakiye"]
    bakiye = Decimal("15250.75")
    d0 = date(2026, 10, 1)
    truth = []
    for i in range(n):
        amt = Decimal(rnd.randint(100, 2500000)) / 100
        borc = rnd.random() < 0.6
        bakiye += -amt if borc else amt
        truth.append({
            "Tarih": (d0 + timedelta(days=i // 3)).strftime("%d.%m.%Y"),
            "Açıklama": rnd.choice(ACIKLAMALAR) + ("" if rnd.random() < .7 else f" REF {rnd.randint(100000, 999999)}"),
            "Borç": tr(amt) if borc else "",
            "Alacak": "" if borc else tr(amt),
            "Bakiye": tr(bakiye),
        })
    total_b = sum(Decimal(t["Borç"].replace(".", "").replace(",", ".")) for t in truth if t["Borç"])
    total_a = sum(Decimal(t["Alacak"].replace(".", "").replace(",", ".")) for t in truth if t["Alacak"])

    if ters:
        truth.reverse()
    if bozuk is not None:
        b = Decimal(truth[bozuk]["Bakiye"].replace(".", "").replace(",", ".")) + 1
        truth[bozuk]["Bakiye"] = tr(b)
    rows = list(truth)
    page, y = 0, 0

    def new_page():
        nonlocal page, y
        if page:
            c.setFont("Arial", 7)
            c.drawCentredString(W / 2, 30, f"Sayfa {page}")
            c.showPage()
        page += 1
        y = H - 50
        c.setFont("ArialB", 12)
        c.drawString(40, y, "ÖRNEK BANKASI A.Ş. — HESAP ÖZETİ")
        y -= 16
        c.setFont("Arial", 8)
        if page == 1:
            c.drawString(40, y, "Müşteri: Ayşe YILMAZ")
            c.drawString(300, y, "IBAN: TR12 0001 2345 6789 0123 4567 89")
            y -= 12
            c.drawString(40, y, "Dönem: 01.10.2026 - 31.10.2026")
            c.drawString(300, y, "Şube: Üsküdar (1234)")
            y -= 20
        c.setFont("ArialB", 8)
        c.drawString(cols["tarih"], y, head[0])
        c.drawString(cols["aciklama"], y, head[1])
        for k, h in zip(("borc", "alacak", "bakiye"), head[2:]):
            c.drawRightString(cols[k], y, h)
        y -= 14

    new_page()
    for r in rows:
        lines = _wrap(r["Açıklama"], "Arial", 8, desc_w)
        if y - 11 * len(lines) < 50:
            new_page()
        c.setFont("Arial", 8)
        c.drawString(cols["tarih"], y, r["Tarih"])
        for k, h in (("borc", "Borç"), ("alacak", "Alacak"), ("bakiye", "Bakiye")):
            if r[h]:
                c.drawRightString(cols[k], y, r[h])
        for j, ln in enumerate(lines):
            c.drawString(cols["aciklama"], y - j * 10, ln)
        y -= 10 * len(lines) + 5
    y -= 6
    c.setFont("ArialB", 8)
    c.drawString(cols["aciklama"], y, "TOPLAM")
    c.drawRightString(cols["borc"], y, tr(total_b))
    c.drawRightString(cols["alacak"], y, tr(total_a))
    c.setFont("Arial", 7)
    c.drawCentredString(W / 2, 30, f"Sayfa {page}")
    c.save()
    return head, truth, {"TOPLAM": [tr(total_b), tr(total_a)]}


def fatura(path: Path):
    """Çizgili tablo: birleşik başlık hücresi, çok satırlı hücre, TL'li tutarlar."""
    c = canvas.Canvas(str(path), pagesize=A4)
    c.setFont("ArialB", 14)
    c.drawString(40, H - 50, "FATURA No: 2026-000123")
    c.setFont("Arial", 9)
    c.drawString(40, H - 66, "Tarih: 05.11.2026")
    xs = [40, 70, 270, 330, 420, 500]
    head = ["Sıra", "Ürün / Hizmet", "Miktar", "Birim Fiyat", "Tutar"]
    items = [
        ["1", "Dizüstü Bilgisayar 15.6\" i7 16GB RAM 512GB SSD", "2", "32.450,00 TL", "64.900,00 TL"],
        ["2", "Kablosuz Mouse", "10", "349,90 TL", "3.499,00 TL"],
        ["3", "Kurulum ve yapılandırma hizmeti (yerinde, 4 saat)", "1", "4.000,00 TL", "4.000,00 TL"],
        ["4", "USB-C Dock", "3", "2.199,99 TL", "6.599,97 TL"],
    ]
    rh = 22
    top = H - 100
    # üst başlık: "Fiyatlandırma" son iki sütunu kaplıyor
    y = top
    c.setFont("ArialB", 9)
    c.rect(xs[0], y - rh, xs[3] - xs[0], rh)
    c.rect(xs[3], y - rh, xs[5] - xs[3], rh)
    c.drawString(xs[0] + 4, y - 15, "Kalemler")
    c.drawString(xs[3] + 4, y - 15, "Fiyatlandırma")
    y -= rh
    for row, font in [(head, "ArialB")] + [(r, "Arial") for r in items]:
        lines_per_cell = [_wrap(t, font, 9, xs[i + 1] - xs[i] - 8) for i, t in enumerate(row)]
        h = rh + 11 * (max(len(l) for l in lines_per_cell) - 1)
        c.setFont(font, 9)
        for i, lines in enumerate(lines_per_cell):
            c.rect(xs[i], y - h, xs[i + 1] - xs[i], h)
            for j, ln in enumerate(lines):
                c.drawString(xs[i] + 4, y - 15 - 11 * j, ln)
        y -= h
    # toplam satırı: ilk 4 sütun birleşik
    c.setFont("ArialB", 9)
    c.rect(xs[0], y - rh, xs[4] - xs[0], rh)
    c.rect(xs[4], y - rh, xs[5] - xs[4], rh)
    c.drawString(xs[0] + 4, y - 15, "GENEL TOPLAM (KDV %20 dahil)")
    c.drawString(xs[4] + 4, y - 15, "78.998,97 TL")
    c.save()
    return head, items


def zor(path: Path):
    """Kenar durumlar. Türkçe harfler bilerek eski 1254 font hatasıyla (ý, þ, ð) basılır."""
    c = canvas.Canvas(str(path), pagesize=A4)
    y = H - 60
    c.setFont("ArialB", 11)
    for dx in (0, 0.2):  # sahte kalın: aynı metin iki kez
        c.drawString(40 + dx, y, "HESAP DÖKÜMÜ")
    y -= 24
    rows = [
        ("Alan", "Deðer"),
        ("Hesap No", "00123456"),
        ("Kart No", "4543 1234 5678 9012"),
        ("Kart No2", "4543123456789012"),
        ("Formül gibi", "=SUM(A1:A3)"),
        ("Sýfýr", "-0,00"),
        ("Tarih saat", "07.11.2026 14:35"),
        ("Oran", "%18"),
        ("Eksi sonda", "1.250,00-"),
        ("Parantez", "(2.000,50)"),
        ("Ýþlem yeri", "Ýstanbul Þiþli Daðýtým"),
        ("Tutar TL", "₺12.345,67"),
    ]
    for a, b in rows:
        c.setFont("Arial", 9)
        c.drawString(40, y, a)
        c.drawString(200, y, b)
        y -= 14
    c.save()
    return rows


def rapor_en(path: Path):
    """İngilizce sayı biçimi, parantezli negatif, yüzde."""
    c = canvas.Canvas(str(path), pagesize=A4)
    c.setFont("ArialB", 12)
    c.drawString(40, H - 50, "Quarterly Results (USD)")
    head = ["Region", "Revenue", "Cost", "Profit", "Margin"]
    data = [
        ["North", "1,250,400.50", "980,100.25", "270,300.25", "21.6%"],
        ["South", "845,000.00", "910,250.75", "(65,250.75)", "-7.7%"],
        ["East", "2,004,300.10", "1,500,000.00", "504,300.10", "25.2%"],
        ["West", "99.99", "0.00", "99.99", "100.0%"],
    ]
    xs = [40, 220, 320, 420, 500]
    y = H - 80
    for row, font in [(head, "ArialB")] + [(r, "Arial") for r in data]:
        c.setFont(font, 9)
        c.drawString(xs[0], y, row[0])
        for x, t in zip(xs[1:], row[1:]):
            c.drawRightString(x + 60, y, t)
        y -= 14
    c.save()
    return head, data

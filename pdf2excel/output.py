"""Excel yazma ve yazılan dosyanın PDF'e karşı doğrulanması."""
from __future__ import annotations

from collections import Counter, defaultdict

from openpyxl import Workbook, load_workbook
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from .checks import balance_chain
from .engine import Result
from .values import render

NOTE_TR = {"tasma": "Metin birden fazla sütuna taşıyor",
           "birlesik": "Aynı sütuna düşen parçalar birleştirildi",
           "devam": "Alt satırdaki devam metni eklendi",
           "sayfa": "Sonraki sayfanın başındaki devam metni eklendi"}


def _sheet_title(name: str) -> str:
    for ch in '[]:*?/\\':
        name = name.replace(ch, " ")
    return name[:31]


def write(result: Result, path: str, report: dict | None = None) -> None:
    wb = Workbook()
    wb.remove(wb.active)
    for s in result.sheets:
        ws = wb.create_sheet(_sheet_title(s.name))
        widths: dict[int, int] = defaultdict(int)
        for ri, row in enumerate(s.rows):
            for c, text in row.cells.items():
                cell = ws.cell(ri + 1, c + 1)
                if (ri, c) in s.typed:
                    cell.value, cell.number_format = s.typed[(ri, c)]
                else:
                    cell.value = ILLEGAL_CHARACTERS_RE.sub("�", text)
                    cell.data_type = "s"          # "=..." formül sanılmasın
                    cell.number_format = "@"
                if row.kind == "ruled":
                    cell.alignment = Alignment(wrap_text=True, vertical="top")
                widths[c] = max(widths[c], min(60, len(text) + 2))
        for r0, c0, r1, c1 in s.merges:
            if (r0, c0) != (r1, c1):
                ws.merge_cells(start_row=r0 + 1, start_column=c0 + 1, end_row=r1 + 1, end_column=c1 + 1)
        for c, w in widths.items():
            ws.column_dimensions[get_column_letter(c + 1)].width = max(8, w)
    if report is not None:
        _report_sheet(wb, report)
    wb.save(path)


def _report_sheet(wb, rep: dict):
    ws = wb.create_sheet("Doğrulama")
    bold = Font(bold=True)
    ok = PatternFill("solid", fgColor="D9F2E3")
    bad = PatternFill("solid", fgColor="F8D7DA")
    rows = [("Sonuç", "BAŞARILI" if rep["ok"] else "SORUN VAR")]
    for p in rep["pages"]:
        rows.append((f"Sayfa {p['page']}", f"{p['pdf_chars']} karakter, "
                     + ("birebir" if p["ok"] else f"fark: {p['diff']}")))
    for b in rep["balance"]:
        rows.append((f"Bakiye zinciri ({b['sheet']})",
                     f"{b['checked']} işlem kontrol edildi, {len(b['errors'])} kopukluk · "
                     f"açılış {b['opening']} → kapanış {b['closing']} · {b['order']}"))
        for e in b["errors"]:
            rows.append(("  Kopukluk", f"{b['sheet']}!{e['ref']}: beklenen {e['expected']}, PDF'te {e['actual']}"))
    if not rep["balance"]:
        rows.append(("Bakiye zinciri", "Uygulanmadı (Bakiye + Tutar ya da Borç/Alacak başlığı bulunamadı)"))
    rows += [("Değer kontrolü", f"{rep['cells']} hücre, {len(rep['value_errors'])} hata"),
             ("Sayı biçimi", rep["locale"] or "belirlenemedi"),
             ("Sayıya çevrilen", rep["info"].get("sayi", 0)),
             ("Tarihe çevrilen", rep["info"].get("tarih", 0)),
             ("Birleştirilen satır", rep["info"].get("birlestirilen_satir", 0)),
             ("Atlanan tekrar satır", f"{rep['info'].get('atlanan_tekrar_satir', 0)} "
              "(her sayfada aynı yerde tekrar eden üst/alt bilgi; ilk sayfada yazıldı)"),
             ("Kontrol önerilen hücre", len(rep["flags"]))]
    for w in rep["warnings"]:
        rows.append(("Uyarı", w))
    for f in rep["flags"]:
        rows.append((f"{f['sheet']}!{f['ref']}", f"{', '.join(NOTE_TR[n] for n in f['notes'])}: {f['text']}"))
    for i, (a, b) in enumerate(rows, 1):
        ws.cell(i, 1, a).font = bold
        ws.cell(i, 2, b)
    ws["B1"].fill = ok if rep["ok"] else bad
    ws.column_dimensions["A"].width = 26
    ws.column_dimensions["B"].width = 90


def verify(result: Result, path: str) -> dict:
    """Dosyayı yeniden açar; her hücrenin Excel'de görünecek metnini kaynakla,
    her sayfanın karakterlerini PDF'in karakterleriyle karşılaştırır."""
    wb = load_workbook(path)
    loc = result.locale or "en"
    value_errors, flags = [], []
    out_chars: dict[int, Counter] = defaultdict(Counter)
    cells = 0
    for s in result.sheets:
        ws = wb[_sheet_title(s.name)]
        expected = set()
        for ri, row in enumerate(s.rows):
            for c, src in row.cells.items():
                cells += 1
                expected.add((ri + 1, c + 1))
                cell = ws.cell(ri + 1, c + 1)
                shown = render(cell.value, cell.number_format, loc)
                ref = f"{get_column_letter(c + 1)}{ri + 1}"
                if shown != src:
                    value_errors.append({"sheet": s.name, "ref": ref, "pdf": src, "excel": shown})
                out_chars[row.page].update(ch for ch in shown if not ch.isspace())
                if row.notes.get(c):
                    flags.append({"sheet": s.name, "ref": ref, "row": ri, "col": c,
                                  "notes": sorted(row.notes[c]), "text": src[:80]})
            for pg, moved in row.moved:  # sayfa geçişinde taşınan metin kendi sayfasına sayılır
                out_chars[row.page] -= moved
                out_chars[pg] += moved
        for r in ws.iter_rows():
            for cell in r:
                if cell.value not in (None, "") and (cell.row, cell.column) not in expected:
                    value_errors.append({"sheet": s.name, "ref": cell.coordinate,
                                         "pdf": "", "excel": str(cell.value)})
    pages = []
    for p, pc in sorted(result.page_chars.items()):
        oc = out_chars.get(p, Counter())
        missing, extra = pc - oc, oc - pc
        diff = ""
        if missing:
            diff += "eksik " + "".join(sorted(missing.elements()))[:40]
        if extra:
            diff += " fazla " + "".join(sorted(extra.elements()))[:40]
        pages.append({"page": p, "pdf_chars": sum(pc.values()), "ok": not missing and not extra,
                      "diff": diff.strip()})
    balance = [b for b in (balance_chain(s, result.locale) for s in result.sheets) if b]
    ok = not value_errors and all(p["ok"] for p in pages) and not any(b["errors"] for b in balance)
    return {"ok": ok, "pages": pages, "cells": cells, "value_errors": value_errors, "balance": balance,
            "flags": flags, "locale": result.locale, "info": result.info, "warnings": result.warnings}


def convert_file(pdf_path: str, xlsx_path: str, opts=None) -> tuple[Result, dict]:
    from .engine import convert
    res = convert(pdf_path, opts)
    write(res, xlsx_path)
    rep = verify(res, xlsx_path)
    write(res, xlsx_path, rep)          # doğrulama sayfasıyla son hali
    rep = verify(res, xlsx_path)        # son dosya da aynı kontrolden geçer
    return res, rep

"""Komut satırı:  python -m pdf2excel dosya.pdf [dosya2.pdf ...] [--sayfa] [--tr|--en]
Her PDF'in yanına aynı adla .xlsx yazar; doğrulama başarısızsa çıkış kodu 1 olur."""
import sys
from pathlib import Path

from . import Options, convert_file


def main(argv):
    files = [a for a in argv if not a.startswith("--")]
    if not files:
        print(__doc__)
        return 2
    opts = Options(mode="sayfa" if "--sayfa" in argv else "tek",
                   locale="tr" if "--tr" in argv else "en" if "--en" in argv else "auto")
    failed = 0
    for f in files:
        p = Path(f)
        for pdf in sorted(p.glob("*.pdf")) if p.is_dir() else [p]:
            out = pdf.with_suffix(".xlsx")
            _, rep = convert_file(str(pdf), str(out), opts)
            flags = sum(1 for x in rep["flags"] if set(x["notes"]) - {"devam"})
            print(f"{'OK   ' if rep['ok'] else 'SORUN'} {pdf.name} -> {out.name} "
                  f"({rep['cells']} hücre, {flags} kontrol önerisi)")
            for b in rep["balance"]:
                print(f"      bakiye zinciri: {b['checked']} işlem, {len(b['errors'])} kopukluk "
                      f"({b['opening']} → {b['closing']})")
            for w in rep["warnings"]:
                print("      uyarı:", w)
            failed += not rep["ok"]
    return 1 if failed else 0


sys.exit(main(sys.argv[1:]))

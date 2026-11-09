"""Windows için PDF2Excel.exe üretir:  .venv\\Scripts\\python.exe build.py
Çıktı: dist\\PDF2Excel.exe"""
import subprocess
import sys
from pathlib import Path

from PIL import Image, ImageDraw

ROOT = Path(__file__).parent
ICON = ROOT / "assets" / "icon.ico"


def make_icon():
    ICON.parent.mkdir(exist_ok=True)
    S = 256
    img = Image.new("RGBA", (S, S), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    d.rounded_rectangle((8, 8, S - 8, S - 8), radius=56, fill=(16, 124, 65, 255))
    for y, x2 in ((88, 196), (128, 196), (168, 148)):
        d.rounded_rectangle((60, y - 10, x2, y + 10), radius=10, fill=(255, 255, 255, 255))
    img.save(ICON, sizes=[(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)])
    img.save(ROOT / "assets" / "icon.png")


if __name__ == "__main__":
    make_icon()
    subprocess.check_call([
        sys.executable, "-m", "PyInstaller", "--noconfirm", "--clean", "--onefile", "--windowed",
        "--name", "PDF2Excel", "--icon", str(ICON),
        "--add-data", f"{ROOT / 'templates'};templates",
        "--collect-data", "pdfminer",
        str(ROOT / "desktop.py"),
    ], cwd=ROOT)
    print("Hazır:", ROOT / "dist" / "PDF2Excel.exe")

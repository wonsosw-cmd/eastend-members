# -*- coding: utf-8 -*-
"""매장별 QR PNG 일괄 생성 (인쇄용 포스터는 docs/qr.html 사용 권장)

사용법:
    py -3.14 -m pip install --user segno
    py -3.14 tools/make_qr.py [출력폴더]
"""
import re
import sys
from pathlib import Path

import segno

BASE_URL = "https://wonsosw-cmd.github.io/eastend-members/"
ROOT = Path(__file__).resolve().parent.parent
OUT = Path(sys.argv[1]) if len(sys.argv) > 1 else ROOT / "qr_png"


def load_stores():
    """docs/stores.js에서 매장 목록 파싱"""
    txt = (ROOT / "docs" / "stores.js").read_text(encoding="utf-8")
    pat = re.compile(r'id:\s*"([^"]+)".*?brandKo:\s*"([^"]+)".*?name:\s*"([^"]+)"')
    return [m.groups() for m in pat.finditer(txt)]


def main():
    OUT.mkdir(parents=True, exist_ok=True)
    stores = load_stores()
    for sid, brand, name in stores:
        url = f"{BASE_URL}?store={sid}"
        qr = segno.make(url, error="m")
        safe = name.replace(" ", "").replace("/", "_")
        path = OUT / f"{sid}_{brand}_{safe}.png"
        qr.save(path, scale=12, border=3)
        print(f"{sid}  {brand} {name}  ->  {path.name}")
    print(f"\n완료: {len(stores)}개 QR 생성 -> {OUT}")


if __name__ == "__main__":
    main()

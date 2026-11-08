#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""build_data.py - sinh data.js cho web bao diem 12A10.

Doc Excel (bang diem tuan + diem giua ki), tinh lai B00 / Tong / xep hang
bang Python (khong tin cong thuc trong file), roi ghi ra data.js.

Cach dung:
    python build_data.py --weekly "duong/dan/Bang diem tuan.xlsx" \
                         --gk1    "duong/dan/GK1.xlsx"

    # chi cap nhat 1 nguon, giu nguyen phan con lai trong data.js hien co
    python build_data.py --gk1 "GK1.xlsx"

    # xuat ra file khac de kiem tra truoc khi ghi de
    python build_data.py --gk1 "GK1.xlsx" --out data_test.js

Yeu cau: pip install openpyxl
"""

from __future__ import annotations

import argparse
import json
import os
import re
import sys
import unicodedata

try:
    import openpyxl
except ImportError:  # pragma: no cover
    sys.exit("Thieu openpyxl. Cai bang: pip install openpyxl")

# 11 mon cua form giua ki, theo dung thu tu cot D..N
GK1_KEYS = ["toan", "ly", "anh", "van", "su", "dia", "hoa", "sinh", "tin", "ktpl", "gdqp"]
GK1_LABELS = {
    "toan": "Toán", "ly": "Lý", "anh": "Anh", "van": "Văn", "su": "Sử",
    "dia": "Địa", "hoa": "Hoá", "sinh": "Sinh", "tin": "Tin",
    "ktpl": "KTPL", "gdqp": "GDQP",
}
WEEKLY_KEYS = ["toan", "hoa", "sinh", "van"]

SUBJECT_MAP = {
    "TOAN": "toan", "LY": "ly", "ANH": "anh", "VAN": "van", "SU": "su",
    "DIA": "dia", "HOA": "hoa", "SINH": "sinh", "TIN": "tin",
    "KTPL": "ktpl", "GDQP": "gdqp",
}


# --------------------------------------------------------------------------- #
# Tien ich
# --------------------------------------------------------------------------- #
def norm(text) -> str:
    """Bo dau, viet hoa, gop khoang trang: 'Hóa' -> 'HOA'."""
    if text is None:
        return ""
    s = unicodedata.normalize("NFD", str(text))
    s = "".join(ch for ch in s if unicodedata.category(ch) != "Mn")
    s = s.replace("đ", "d").replace("Đ", "D")
    s = re.sub(r"\s+", " ", s)
    return s.strip().upper()


def r2(value):
    """Lam tron 2 chu so thap phan, tra None neu khong phai so."""
    if value is None or isinstance(value, bool):
        return None
    try:
        n = float(value)
    except (TypeError, ValueError):
        return None
    if n != n:  # NaN
        return None
    return round(n, 2)


def rank_map(rows, key):
    """Xep hang kieu Excel RANK(...,0): dong hang chia cung thu hang."""
    valid = [r for r in rows if r.get(key) is not None]
    out = {}
    for r in valid:
        out[r["sbd"]] = 1 + sum(1 for o in valid if o[key] > r[key])
    return out


def to_sbd(value):
    if value is None:
        return None
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    s = str(value).strip()
    if re.fullmatch(r"\d+\.0", s):
        s = s[:-2]
    return s or None


# --------------------------------------------------------------------------- #
# Doc sheet
# --------------------------------------------------------------------------- #
def header_map(ws, max_scan: int = 12):
    """Tim dong tieu de va tra ve (row_idx, {cot: khoa})."""
    best = None
    for r in range(1, min(ws.max_row, max_scan) + 1):
        cmap = {}
        for c in range(1, ws.max_column + 1):
            raw = ws.cell(r, c).value
            h = norm(raw)
            if not h:
                continue
            key = classify_header(h)
            if key and key not in cmap:
                cmap[key] = c
        if "sbd" in cmap and len(cmap) >= 4:
            best = (r, cmap)
            if len(cmap) >= 6:
                break
    if not best:
        raise ValueError("Khong tim thay dong tieu de co cot SBD")
    return best


def classify_header(h: str):
    if h == "SBD":
        return "sbd"
    if h in ("HO", "HO VA"):
        return "ho"
    if h == "TEN":
        return "ten"
    if "HO" in h and "TEN" in h:
        return "hoten"
    if "B00" in h:
        return "b00"
    if h.startswith("XH") and "KHOI" in h:
        return "xhb"
    if h.startswith("XH"):
        return "xh"
    if h.startswith("TONG"):
        return "tong"
    return SUBJECT_MAP.get(h)


def read_sheet(ws, keys):
    """Doc mot sheet thanh list dict {sbd, hoten, <mon>...}."""
    hrow, cmap = header_map(ws)
    rows = []
    for r in range(hrow + 1, ws.max_row + 1):
        sbd = to_sbd(ws.cell(r, cmap["sbd"]).value)
        if not sbd or not re.fullmatch(r"\d{3,6}", sbd):
            continue
        if "hoten" in cmap:
            hoten = str(ws.cell(r, cmap["hoten"]).value or "").strip()
        else:
            ho = str(ws.cell(r, cmap["ho"]).value or "").strip() if "ho" in cmap else ""
            ten = str(ws.cell(r, cmap["ten"]).value or "").strip() if "ten" in cmap else ""
            hoten = (ho + " " + ten).strip()
        rec = {"sbd": sbd, "hoten": re.sub(r"\s+", " ", hoten)}
        for key in keys:
            col = cmap.get(key)
            rec[key] = r2(ws.cell(r, col).value) if col else None
        rows.append(rec)
    return rows


def read_weekly(path):
    """Doc file bang diem tuan: moi sheet 'Lan N' -> mot dot."""
    wb = openpyxl.load_workbook(path, data_only=True)
    out = {}
    for ws in wb.worksheets:
        m = re.search(r"L[ẦA]N\s*(\d+)", norm(ws.title), re.IGNORECASE)
        if not m:
            continue
        lan = "Lần " + m.group(1)
        rows = read_sheet(ws, WEEKLY_KEYS)
        for r in rows:
            vals = [r[k] for k in WEEKLY_KEYS]
            if all(v is None or v == 0 for v in vals):
                for k in WEEKLY_KEYS:
                    r[k] = None
            if all(r[k] is not None for k in ("toan", "hoa", "sinh")):
                r["b00"] = round(r["toan"] + r["hoa"] + r["sinh"], 2)
            else:
                r["b00"] = None
        xh = rank_map(rows, "b00")
        for r in rows:
            r["xh"] = xh.get(r["sbd"])
        rows.sort(key=lambda x: x["sbd"])
        out[lan] = rows
    if not out:
        raise ValueError("Khong thay sheet nao ten 'Lan N' trong file: " + path)
    return out


def read_gk1(path):
    """Doc file diem giua ki (11 mon)."""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.worksheets[0]
    rows = read_sheet(ws, GK1_KEYS)
    if not rows:
        raise ValueError("Khong doc duoc hoc sinh nao trong file GK1: " + path)
    for r in rows:
        present = [r[k] for k in GK1_KEYS if r[k] is not None]
        if present and all(v == 0 for v in present) and len(present) == len(GK1_KEYS):
            for k in GK1_KEYS:
                r[k] = None
            present = []
        if all(r[k] is not None for k in ("toan", "hoa", "sinh")):
            r["b00"] = round(r["toan"] + r["hoa"] + r["sinh"], 2)
        else:
            r["b00"] = None
        r["tong"] = round(sum(present), 2) if present else None
        r["so_mon"] = len(present)
    xhb = rank_map(rows, "b00")
    xh = rank_map(rows, "tong")
    for r in rows:
        r["xhb"] = xhb.get(r["sbd"])
        r["xh"] = xh.get(r["sbd"])
    rows.sort(key=lambda x: x["sbd"])
    return rows


def has_scores_gk1(rows) -> bool:
    return any(r.get(k) is not None for r in rows for k in GK1_KEYS)


# --------------------------------------------------------------------------- #
# data.js
# --------------------------------------------------------------------------- #
def load_existing(path):
    """Doc lai DATA tu data.js neu co (dinh dang JSON sau dau '=')."""
    if not os.path.exists(path):
        return {}
    raw = open(path, "r", encoding="utf-8").read()
    m = re.search(r"const\s+DATA\s*=\s*(\{.*\})\s*;", raw, re.S)
    if not m:
        return {}
    try:
        return json.loads(m.group(1))
    except json.JSONDecodeError:
        print("! data.js hien co khong doc duoc - se tao moi tu Excel", file=sys.stderr)
        return {}


def dump_js(data: dict) -> str:
    lines = ["// DIEM 12A10 - sinh tu build_data.py (dung sua tay)",
             "// XH chi hien tren WEB, KHONG dua vao tin nhan gui PH",
             "const DATA = {"]
    keys = list(data.keys())
    for i, key in enumerate(keys):
        body = json.dumps(data[key], ensure_ascii=False)
        comma = "," if i < len(keys) - 1 else ""
        lines.append('"%s": %s%s' % (key, body, comma))
    lines.append("};")
    siso = len(data.get(keys[-1], [])) if keys else 0
    lines.append("const SISO = %d;" % siso)
    lines.append("")
    return "\n".join(lines)


def bump_cache_busters(folder, stamp=None):
    """Tang ?v= cua data.js trong index.html / gk1.html de trinh duyet tai ban moi."""
    import datetime
    stamp = stamp or datetime.datetime.now().strftime("%Y%m%d_%H%M")
    changed = []
    for name in ("index.html", "gk1.html"):
        p = os.path.join(folder, name)
        if not os.path.exists(p):
            continue
        with open(p, "r", encoding="utf-8", newline="") as f:
            raw = f.read()
        new = re.sub(r"data\.js\?v=[0-9A-Za-z_\-]+", "data.js?v=" + stamp, raw)
        if new != raw:
            with open(p, "w", encoding="utf-8", newline="") as f:
                f.write(new)
            changed.append(name)
    return changed, stamp


def main(argv=None):
    ap = argparse.ArgumentParser(description="Sinh data.js cho web bao diem 12A10")
    ap.add_argument("--weekly", help="File Excel bang diem tuan (nhieu sheet Lan N)")
    ap.add_argument("--gk1", help="File Excel diem giua ki 1 (11 mon)")
    ap.add_argument("--out", default="data.js", help="File JS dau ra (mac dinh data.js)")
    ap.add_argument("--base", default=None, help="data.js cu de giu cac dot khong co trong Excel")
    ap.add_argument("--force", action="store_true",
                    help="Ghi de ca khi ten hoc sinh lech so voi data.js hien co")
    ap.add_argument("--no-bump", action="store_true",
                    help="Khong tu tang ?v= cua data.js trong index.html / gk1.html")
    args = ap.parse_args(argv)

    out_path = args.out
    base_path = args.base or out_path
    data = load_existing(base_path)

    if args.weekly:
        weekly = read_weekly(args.weekly)
        for lan, rows in weekly.items():
            old = {r["sbd"]: r.get("hoten", "") for r in data.get(lan, [])}
            lech = [
                (r["sbd"], old[r["sbd"]], r["hoten"])
                for r in rows
                if r["sbd"] in old and norm(old[r["sbd"]]) != norm(r["hoten"])
            ]
            if lech and not args.force:
                print("! %s: ten lech o %d SBD so voi data.js -> BO QUA (them --force neu chac chan)"
                      % (lan, len(lech)))
                for sbd, cu, moi in lech[:5]:
                    print("    %s: data.js='%s' | Excel='%s'" % (sbd, cu, moi))
                continue
            data[lan] = rows
            print("+ %s: %d hoc sinh" % (lan, len(rows)))

    if args.gk1:
        rows = read_gk1(args.gk1)
        if has_scores_gk1(rows):
            data["Giữa kì 1"] = rows
            filled = sum(1 for r in rows if r["so_mon"] > 0)
            print("+ Giữa kì 1: %d hoc sinh, %d em da co diem" % (len(rows), filled))
        else:
            data.pop("Giữa kì 1", None)
            print("! File GK1 chua co diem - khong dua 'Giữa kì 1' len web")

    if not data:
        sys.exit("Khong co du lieu nao de ghi.")

    # Sap xep: Lan 1..N truoc, Giua ki 1 sau cung
    def sort_key(k):
        m = re.fullmatch(r"Lần (\d+)", k)
        return (0, int(m.group(1))) if m else (1, k)

    ordered = {k: data[k] for k in sorted(data.keys(), key=sort_key)}
    js = dump_js(ordered)

    with open(out_path, "w", encoding="utf-8", newline="\n") as f:
        f.write(js)
    print("Da ghi:", os.path.abspath(out_path), "(%d dot, %.1f KB)" % (len(ordered), len(js) / 1024))

    # Tu tang ?v= de trinh duyet tai data.js moi (chi khi ghi dung file data.js)
    if not args.no_bump and os.path.basename(out_path) == "data.js":
        changed, stamp = bump_cache_busters(os.path.dirname(os.path.abspath(out_path)) or ".")
        if changed:
            print("Da tang cache-buster data.js?v=%s trong: %s" % (stamp, ", ".join(changed)))
        else:
            print("! Khong tim thay data.js?v=... de tang trong index.html / gk1.html")


if __name__ == "__main__":
    main()

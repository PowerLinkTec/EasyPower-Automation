"""
dez_inspect.py — read EasyPower .dez files without EasyPower.

A .dez file is a ZIP archive. The member `EZP.U` holds the model as a series of
RECORD ... ENDREC blocks of `key=value` lines (CRLF, tab-indented). This module
parses that into Python so you can inventory and diff your 36 scenario files
before spending time on the (slow, UI-driven) report run.

It does NOT compute results — short-circuit/arc-flash/coordination numbers come
from EasyPower's solver and are not stored in the input file.

CLI:
    python dez_inspect.py path\\to\\file.dez            # summarize one file
    python dez_inspect.py path\\to\\folder --manifest manifest.csv   # all .dez in folder
"""

import csv
import sys
import zipfile
import struct
import argparse
from pathlib import Path


def _decode_value(v: str):
    """EasyPower encodes numbers as reinterpreted bit patterns.
    'x<hex>'  -> float32 from those 4 bytes
    'L<int>'  -> float64 from that 64-bit pattern (often) or a plain long
    Returns a float when decodable, else the original string."""
    if not v:
        return v
    try:
        if v[0] == "x":
            bits = int(v[1:], 16)
            return struct.unpack("<f", struct.pack("<I", bits & 0xFFFFFFFF))[0]
        if v[0] == "L":
            n = int(v[1:])
            # Heuristic: large magnitudes are double bit patterns; small are ints.
            if abs(n) > (1 << 40):
                return struct.unpack("<d", struct.pack("<q", n))[0]
            return n
    except (ValueError, struct.error):
        pass
    return v


def parse_dez(path: Path) -> dict:
    """Return {'header': {...}, 'records': [ {field: value, ...}, ... ]}."""
    with zipfile.ZipFile(path) as z:
        raw = z.read("EZP.U").decode("utf-8-sig", errors="replace")

    header, records = {}, []
    cur, in_header = None, False
    for line in raw.splitlines():
        if line.startswith("HEADER"):
            in_header, cur = True, header
            continue
        if line.startswith("RECORD"):
            in_header = False
            cur = {}
            records.append(cur)
            continue
        if line.startswith("ENDREC") or line.startswith("EOF"):
            cur, in_header = None, False
            continue
        if cur is not None and "=" in line:
            k, _, val = line.strip().partition("=")
            cur[k] = val
    return {"header": header, "records": records}


def summarize(path: Path) -> dict:
    """One-row summary suitable for a manifest."""
    model = parse_dez(path)
    recs = model["records"]
    named = [r.get("szName", "") for r in recs if r.get("szName")]
    # Drop auto-generated connection-node names (26-char all-caps/digits tokens).
    real = [n for n in named if not (len(n) == 26 and n.isalnum() and n.isupper())]
    base_mva = _decode_value(model["header"].get("fBaseMva", ""))
    return {
        "file": path.name,
        "version": model["header"].get("nVersion", ""),
        "frequency_hz": model["header"].get("nSysFrequency", ""),
        "base_mva": round(base_mva, 4) if isinstance(base_mva, float) else base_mva,
        "record_count": len(recs),
        "named_equipment": len(real),
        "equipment_sample": "; ".join(real[:12]),
    }


def folder_manifest(folder: Path, out_csv: Path):
    files = sorted(folder.glob("*.dez"))
    if not files:
        print(f"No .dez files in {folder}", file=sys.stderr)
        return
    rows = []
    for f in files:
        try:
            rows.append(summarize(f))
        except Exception as e:  # keep going; flag the bad file
            rows.append({"file": f.name, "version": "ERROR", "frequency_hz": "",
                         "base_mva": "", "record_count": "", "named_equipment": "",
                         "equipment_sample": str(e)})
    with out_csv.open("w", newline="", encoding="utf-8") as fh:
        w = csv.DictWriter(fh, fieldnames=list(rows[0].keys()))
        w.writeheader()
        w.writerows(rows)
    print(f"Wrote {out_csv} ({len(rows)} files).")


def main():
    ap = argparse.ArgumentParser(description="Inspect EasyPower .dez files.")
    ap.add_argument("target", type=Path, help=".dez file or a folder of them")
    ap.add_argument("--manifest", type=Path, help="write a CSV manifest (folder mode)")
    args = ap.parse_args()

    if args.target.is_dir():
        folder_manifest(args.target, args.manifest or args.target / "manifest.csv")
    else:
        s = summarize(args.target)
        width = max(len(k) for k in s)
        for k, v in s.items():
            print(f"{k:<{width}} : {v}")


if __name__ == "__main__":
    main()

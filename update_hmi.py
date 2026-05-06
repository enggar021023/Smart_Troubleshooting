import sqlite3
import re
import sys
import argparse
from pathlib import Path
from collections import defaultdict

try:
    import openpyxl
except ImportError:
    print("pip install openpyxl"); sys.exit(1)

DEFAULT_DB    = "data/ammonia.db"
DEFAULT_EXCEL = "data/HMI.xlsx"

HMI_TO_DB_PREFIX = {
    # Flow
    'FI':   ['FT'],              # Flow Indicator → Flow Transmitter
    'FC':   ['FY', 'FT'],        # Flow Controller → FY output / FT
    'FCA':  ['FY', 'FT'],        # Flow Controller Auto → FY/FT
    'FQI':  ['FT'],              # Flow Quantity Indicator → FT

    # Temperature
    'TA':   ['TI', 'TE'],        # Temp Average → TI indicator / TE element
    'TI':   ['TE', 'TW'],        # Temp Indicator → TE element / TW thermowell
    'TC':   ['TY', 'TV', 'TIC'], # Temp Controller → TY/TV output / TIC
    'TCA':  ['TY', 'TV', 'TIC'],

    # Pressure
    'PA':   ['PI', 'PT'],        # Pressure Average → PI / PT transmitter
    'PI':   ['PT'],              # Pressure Indicator → PT transmitter
    'PC':   ['PY', 'PIC'],       # Pressure Controller → PY/PIC
    'PCA':  ['PY', 'PIC', 'PT'], # Pressure Controller Auto → PY/PIC/PT
    'PDA':  ['PDI', 'PDT'],      # ΔP Average → PDI / PDT transmitter
    'PDI':  ['PDT'],             # ΔP Indicator → PDT

    # Level
    'LI':   ['LT'],              # Level Indicator → LT transmitter
    'LC':   ['LT', 'ZLC'],       # Level Controller → LT / ZLC
    'LCA':  ['LV', 'LT'],        # Level Controller Auto → LV/LT

    # Hand/Manual
    'HC':   ['HV'],              # Hand Controller → Hand Valve

    # Vibration/Speed/Position
    'XI':   ['XT'],              # Speed/Vibration Indicator → XT transmitter
    'ZI':   ['XT', 'ZT'],        # Position Indicator → XT/ZT transmitter
    'SI':   ['SE', 'ST'],        # Speed Indicator → SE element

    # Analyzer
    'LA':   ['AI', 'AE'],        # Level/Analyzer → AI indicator / AE element
}

# ═══════════════════════════════════════════════════════════
#  NORMALISASI
# ═══════════════════════════════════════════════════════════
def normalize(tag: str) -> str:
    """Hapus spasi/strip/underscore/titik → uppercase."""
    if not tag:
        return ""
    return re.sub(r"[\s\-_.]", "", str(tag)).upper()


# ═══════════════════════════════════════════════════════════
#  BACA EXCEL
# ═══════════════════════════════════════════════════════════
def read_excel(excel_path: str) -> list[dict]:
    path = Path(excel_path)
    if not path.exists():
        print(f"❌ Excel tidak ditemukan: {excel_path}"); sys.exit(1)

    wb   = openpyxl.load_workbook(str(path), read_only=True, data_only=True)
    ws   = wb.active
    rows = list(ws.iter_rows(values_only=True))
    if len(rows) < 5:
        print("❌ File Excel minimal 5 baris."); sys.exit(1)

    row_pages = rows[1]
    row_links = rows[2]
    num_cols  = max(len(row_pages), len(row_links))
    pages     = []

    for col_idx in range(1, num_cols):
        page_name   = row_pages[col_idx] if col_idx < len(row_pages) else None
        gdrive_link = row_links[col_idx]  if col_idx < len(row_links)  else None
        if not page_name and not gdrive_link:
            continue

        page_name   = str(page_name).strip()   if page_name   else f"COL_{col_idx}"
        gdrive_link = str(gdrive_link).strip()  if gdrive_link else ""

        tags_raw = []
        for row in rows[4:]:
            val = row[col_idx] if col_idx < len(row) else None
            if val and str(val).strip():
                tags_raw.append(str(val).strip())

        pages.append({"page": page_name, "link": gdrive_link, "tags_raw": tags_raw})

    wb.close()
    return pages


# ═══════════════════════════════════════════════════════════
#  BACA DATABASE
# ═══════════════════════════════════════════════════════════
def load_db_tags(db_path: str) -> dict[str, str]:
    """Return {normalized → original_tag_number}"""
    if not Path(db_path).exists():
        print(f"❌ DB tidak ditemukan: {db_path}"); sys.exit(1)

    conn = sqlite3.connect(str(db_path))
    cur  = conn.cursor()
    try:
        cur.execute("SELECT tag_number FROM equipment")
    except sqlite3.OperationalError as e:
        print(f"❌ Error: {e}"); conn.close(); sys.exit(1)

    result = {}
    for (tag,) in cur.fetchall():
        if tag:
            n = normalize(tag)
            if n:
                result[n] = tag
    conn.close()
    return result


# ═══════════════════════════════════════════════════════════
#  MATCHING
# ═══════════════════════════════════════════════════════════
def try_match(raw_tag: str, db_tags: dict[str, str]) -> tuple[str | None, str]:
    """
    Coba match tag HMI ke DB dengan 2 strategi.
    Return: (original_tag_in_db | None, metode)
    """
    n = normalize(raw_tag)

    # Skip tag yang tidak relevan
    if not n:                           return None, "kosong"
    if re.match(r"^\d", n):            return None, "awalan_angka"
    if "DAILY" in n:                   return None, "suffix_daily"

    # ── Strategi 1: Direct match ──────────────────────────────
    if n in db_tags:
        return db_tags[n], "direct"

    # ── Strategi 2: HMI prefix → DB prefix mapping ───────────
    prefix_m = re.match(r"^([A-Z]+)", n)
    if prefix_m:
        hmi_prefix = prefix_m.group(1)
        number_part = n[len(hmi_prefix):]  # angka + suffix (misal: 1042B)

        if hmi_prefix in HMI_TO_DB_PREFIX:
            for db_prefix in HMI_TO_DB_PREFIX[hmi_prefix]:
                candidate = db_prefix + number_part
                if candidate in db_tags:
                    return db_tags[candidate], f"{hmi_prefix}→{db_prefix}"

    return None, "tidak_ditemukan"


def build_updates(pages: list[dict], db_tags: dict[str, str],
                  verbose: bool = False) -> tuple[list, list]:
    update_map: dict[str, str] = {}   # db_tag → link
    unmatched  = []
    method_count = defaultdict(int)

    for page in pages:
        link      = page["link"]
        page_name = page["page"]
        if not link:
            continue

        for raw_tag in page["tags_raw"]:
            db_tag, method = try_match(raw_tag, db_tags)
            if db_tag:
                update_map[db_tag] = link
                method_count[method] += 1
            else:
                if method == "tidak_ditemukan":
                    unmatched.append((raw_tag, page_name))

    if verbose:
        print("\n  Breakdown strategi matching:")
        for m, c in sorted(method_count.items(), key=lambda x: -x[1]):
            print(f"    {m:<20}: {c}")

    return list(update_map.items()), unmatched


# ═══════════════════════════════════════════════════════════
#  UPDATE & VERIFIKASI
# ═══════════════════════════════════════════════════════════
def apply_updates(db_path: str, updates: list, reset: bool = False) -> int:
    conn = sqlite3.connect(str(db_path))
    cur  = conn.cursor()
    if reset:
        cur.execute("UPDATE equipment SET hmi_link = NULL")
        print(f"  🔄 Reset {cur.rowcount} baris.")
    count = 0
    for db_tag, link in updates:
        cur.execute("UPDATE equipment SET hmi_link = ? WHERE tag_number = ?", (link, db_tag))
        count += cur.rowcount
    conn.commit()
    conn.close()
    return count

def verify_update(db_path: str) -> dict:
    conn = sqlite3.connect(str(db_path))
    cur  = conn.cursor()
    cur.execute("SELECT COUNT(*) FROM equipment"); total = cur.fetchone()[0]
    cur.execute("SELECT COUNT(*) FROM equipment WHERE hmi_link IS NOT NULL AND hmi_link != ''")
    filled = cur.fetchone()[0]
    cur.execute("SELECT COUNT(*) FROM equipment WHERE hmi_link IS NULL OR hmi_link = ''")
    empty = cur.fetchone()[0]
    cur.execute("SELECT tag_number, hmi_link FROM equipment WHERE hmi_link IS NOT NULL LIMIT 5")
    samples = cur.fetchall()
    conn.close()
    return {"total": total, "filled": filled, "empty": empty, "samples": samples}


# ═══════════════════════════════════════════════════════════
#  MODE DIAGNOSA
# ═══════════════════════════════════════════════════════════
def run_diagnose(pages: list[dict], db_tags: dict[str, str]):
    print("\n" + "═"*64)
    print("  🔬 MODE DIAGNOSA")
    print("═"*64)

    total = sum(len(p["tags_raw"]) for p in pages)
    updates, unmatched = build_updates(pages, db_tags, verbose=True)

    print(f"\n  Total tag Excel   : {total}")
    print(f"  ✅ Matched        : {len(updates)}")
    print(f"  ❌ Tidak match    : {len(unmatched)}")
    if total:
        print(f"  📈 Match rate     : {len(updates)/total*100:.1f}%")

    # Prefix unmatched
    unm_prefix = defaultdict(list)
    for raw, page in unmatched:
        m = re.match(r"^([A-Z]+)", normalize(raw))
        if m: unm_prefix[m.group(1)].append((raw, page))

    print(f"\n─── Prefix yang masih tidak match ───")
    for pref, items in sorted(unm_prefix.items(), key=lambda x: -len(x[1]))[:20]:
        in_db = any(k.startswith(pref) for k in db_tags)
        status = "✅ ada di DB" if in_db else "❌ tidak ada di DB"
        print(f"  {pref:<10}: {len(items):>3} tag  [{status}]")
        for r, p in items[:2]:
            print(f"    ex: '{r}'  (page: {p})")

    print(f"""
─── Penjelasan Tag yang Masih Tidak Match ───────────────────

  PCA / LCA / TCA / FCA  = Controller dalam mode Auto
    → Tag ini adalah STATUS/MODE di DCS, bukan field instrument.
    → Tidak ada di instrument index = WAJAR, tidak perlu di-update.

  FQ / FQI  = Flow Quantity (Totalizer)
    → Tag kalkulasi di DCS historian, bukan physical instrument.
    → Tidak ada di DB = WAJAR.

  GV  = Gate Valve / Group Valve
    → Bukan instrument, tapi equipment. Tidak di instrument index.

  FI / TI / PI / HC yang masih tidak match
    → Angka loop-nya tidak ada di DB sama sekali.
    → Kemungkinan: area/unit yang belum di-input ke instrument index.
    → Cek dengan: python investigate_tags.py --export

─── Match Rate Progression ──────────────────────────────────

  v1 (hanya normalize)          : 437 / 1657  =  26.4%
  v3 (+ prefix mapping HMI→DB)  : 643 / 1610  =  39.9%
  Ceiling realistis              : ~650-700   =  ~42%
    (sisa = tag DCS internal yang memang tidak punya field instrument)
""")


# ═══════════════════════════════════════════════════════════
#  MAIN
# ═══════════════════════════════════════════════════════════
def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--db",      default=DEFAULT_DB)
    parser.add_argument("--excel",   default=DEFAULT_EXCEL)
    parser.add_argument("--dry-run", action="store_true")
    parser.add_argument("--reset",   action="store_true")
    parser.add_argument("--diagnose",action="store_true")
    args = parser.parse_args()

    print("="*64)
    print("  🔧 HMI Link Updater v3 — PT Petrokimia Gresik Ammonia Plant")
    print("="*64)
    print(f"  DB    : {args.db}")
    print(f"  Excel : {args.excel}")
    if args.dry_run:   print("  Mode  : DRY-RUN (tidak ada perubahan)")
    if args.reset:     print("  Reset : ON")
    if args.diagnose:  print("  Mode  : DIAGNOSA")

    print("\n📖 Membaca Excel...")
    pages = read_excel(args.excel)
    total = sum(len(p["tags_raw"]) for p in pages)
    print(f"   {len(pages)} page, {total} tag")

    print("\n🗄️  Membaca database...")
    db_tags = load_db_tags(args.db)
    print(f"   {len(db_tags)} tag di database")

    if args.diagnose:
        run_diagnose(pages, db_tags)
        return

    print("\n🔍 Mencocokkan tag...")
    updates, unmatched = build_updates(pages, db_tags, verbose=True)
    matched = len(updates)
    pct = matched / total * 100 if total else 0

    print(f"\n{'─'*64}")
    print(f"  ✅ Berhasil dicocokkan : {matched:>5}")
    print(f"  ❌ Tidak ditemukan     : {len(unmatched):>5}")
    print(f"  📊 Match rate          : {pct:.1f}%")
    print(f"{'─'*64}")

    if matched == 0:
        print("\n⚠️  Tidak ada match. Jalankan --diagnose."); return

    if args.dry_run:
        print("\n[DRY-RUN] Selesai. Hapus --dry-run untuk update sesungguhnya.")
        return

    print()
    if args.reset:
        print("⚠️  Mode RESET aktif — semua hmi_link akan dikosongkan dulu!")
    confirm = input(f"Update {matched} baris hmi_link di database? [y/N] ").strip().lower()
    if confirm != "y":
        print("❌ Dibatalkan."); return

    print("\n💾 Melakukan update...")
    updated = apply_updates(args.db, updates, reset=args.reset)
    print(f"   {updated} baris berhasil diupdate.")

    print("\n🔎 Verifikasi...")
    stats = verify_update(args.db)
    print(f"   Total : {stats['total']}  |  Terisi: {stats['filled']}  |  Kosong: {stats['empty']}")

    if stats["samples"]:
        print("\n   Sample hasil:")
        for tag, link in stats["samples"]:
            lp = (link[:52] + "...") if link and len(link) > 52 else link
            print(f"     {tag:<22} → {lp}")

    if unmatched:
        log = Path(args.db).parent / "hmi_unmatched_log.txt"
        with open(log, "w", encoding="utf-8") as f:
            f.write("TAG_EXCEL\tPAGE\n")
            for t, p in sorted(unmatched):
                f.write(f"{t}\t{p}\n")
        print(f"\n📝 Log tag tidak match: {log}")

    print("\n✅ Selesai!")
    print("   Jalankan --diagnose untuk melihat breakdown lengkap.")

if __name__ == "__main__":
    main()
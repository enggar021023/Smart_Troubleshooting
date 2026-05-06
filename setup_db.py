import sqlite3
import os
import re
import sys

# ─────────────────────────────────────────────
# KONFIGURASI
# ─────────────────────────────────────────────
DB_PATH    = "data/ammonia.db"
EXCEL_PATH = "data/Ammonia Instrument Index.xlsx"

# ─────────────────────────────────────────────
# 1. BUAT FOLDER & TABEL
# ─────────────────────────────────────────────
def create_database(reset: bool = False):
    os.makedirs("data", exist_ok=True)
    conn   = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()

    if reset:
        cursor.execute("DROP TABLE IF EXISTS equipment")
        print("Tabel lama dihapus.")

    cursor.executescript('''
        CREATE TABLE IF NOT EXISTS equipment (
            Tag_Number          TEXT PRIMARY KEY,
            PID_Number          TEXT DEFAULT '-',
            Service_Description TEXT DEFAULT '-',
            PID_Link            TEXT DEFAULT '-',
            Logic_Diagram       TEXT DEFAULT '-',
            HMI_Link            TEXT DEFAULT '-'
        );
        CREATE INDEX IF NOT EXISTS idx_tag ON equipment(Tag_Number);
    ''')
    conn.commit()
    conn.close()
    print("Tabel equipment siap.")

# ─────────────────────────────────────────────
# 2. IMPORT DARI EXCEL
# ─────────────────────────────────────────────
def clean_col(name: str) -> str:
    return re.sub(r'\s+', ' ', name.replace('\n', ' ')).strip()

def import_from_excel():
    if not os.path.exists(EXCEL_PATH):
        print(f"File Excel tidak ditemukan di '{EXCEL_PATH}'. Lewati import.")
        return 0

    try:
        import pandas as pd
    except ImportError:
        print("pandas tidak terinstall. Jalankan: pip install pandas openpyxl")
        return 0

    try:
        print(f"Membaca file: {EXCEL_PATH}")
        df = pd.read_excel(EXCEL_PATH)
        df.columns = [clean_col(c) for c in df.columns]
        print(f"Kolom Excel terdeteksi: {list(df.columns)}")

        # Mapping kolom Excel ke kolom database
        col_map = {
            'Instrument Tag Number': 'Tag_Number',
            'P&ID Number':           'PID_Number',
            'Service Description':   'Service_Description',
            'LINK P&ID':             'PID_Link',
            'Link P&ID':             'PID_Link',
            'LINK INTERLOCK':        'Logic_Diagram',
            'Link Logic Diagram':    'Logic_Diagram',
            'LINK HMI':              'HMI_Link',
            'Link HMI':              'HMI_Link',
        }
        df = df.rename(columns={k: v for k, v in col_map.items() if k in df.columns})

        if 'Tag_Number' not in df.columns:
            print("Kolom 'Instrument Tag Number' tidak ditemukan di Excel!")
            return 0

        # Bersihkan data
        df['Tag_Number'] = df['Tag_Number'].astype(str).str.strip()
        df = df[df['Tag_Number'].str.lower() != 'nan']
        df = df.drop_duplicates(subset=['Tag_Number'], keep='first')
        df = df.fillna('-')

        # Pastikan semua kolom ada
        for col in ['PID_Number', 'Service_Description', 'PID_Link', 'Logic_Diagram', 'HMI_Link']:
            if col not in df.columns:
                df[col] = '-'

        print(f"Total data siap diimport: {len(df)} baris")
        print(df[['Tag_Number', 'PID_Number', 'Service_Description']].head().to_string(index=False))

        conn   = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()

        inserted = 0
        skipped  = 0
        for _, row in df.iterrows():
            try:
                cursor.execute(
                    '''INSERT OR REPLACE INTO equipment
                        (Tag_Number, PID_Number, Service_Description, PID_Link, Logic_Diagram, HMI_Link)
                       VALUES (?, ?, ?, ?, ?, ?)''',
                    (
                        str(row['Tag_Number']),
                        str(row.get('PID_Number', '-')),
                        str(row.get('Service_Description', '-')),
                        str(row.get('PID_Link', '-')),
                        str(row.get('Logic_Diagram', '-')),
                        str(row.get('HMI_Link', '-')),
                    )
                )
                inserted += 1
            except Exception as e:
                print(f"Gagal insert '{row.get('Tag_Number', '?')}': {e}")
                skipped += 1

        conn.commit()
        conn.close()
        print(f"Import Excel selesai — {inserted} berhasil, {skipped} gagal.")
        return inserted

    except Exception as e:
        print(f"Import Excel gagal: {e}")
        return 0

# ─────────────────────────────────────────────
# 3. VERIFIKASI DATABASE
# ─────────────────────────────────────────────
def verify():
    if not os.path.exists(DB_PATH):
        print("Database belum dibuat.")
        return

    conn   = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()

    cursor.execute("SELECT COUNT(*) FROM equipment")
    total = cursor.fetchone()[0]
    print(f"\nTotal data di database : {total} equipment")

    cursor.execute("SELECT COUNT(*) FROM equipment WHERE PID_Link != '-' AND PID_Link != ''")
    with_link = cursor.fetchone()[0]
    print(f"Equipment dengan PID link    : {with_link}")

    cursor.execute("SELECT COUNT(*) FROM equipment WHERE HMI_Link != '-' AND HMI_Link != ''")
    with_hmi = cursor.fetchone()[0]
    print(f"Equipment dengan HMI link    : {with_hmi}")

    cursor.execute("SELECT COUNT(*) FROM equipment WHERE Logic_Diagram != '-' AND Logic_Diagram != ''")
    with_ld = cursor.fetchone()[0]
    print(f"Equipment dengan Logic Diagram: {with_ld}")

    print(f"\nContoh 5 data pertama:")
    cursor.execute('SELECT Tag_Number, PID_Number, PID_Link FROM equipment LIMIT 5')
    for r in cursor.fetchall():
        link_status = "ada link" if r[2] and r[2] != "-" else "belum ada link"
        print(f"   {r[0]:15s} | {r[1]:15s} | {link_status}")

    conn.close()

# ─────────────────────────────────────────────
# HELPER FUNCTIONS (dipakai main.py & backend.py)
# ─────────────────────────────────────────────
def get_connection():
    return sqlite3.connect(DB_PATH, check_same_thread=False)


def _to_dict(row: tuple) -> dict:
    return {
        'tag_number':          row[0],
        'pid_number':          row[1],
        'service_description': row[2],
        'pid_link':            row[3],
        'logic_diagram':       row[4],
        'hmi_link':            row[5],
    }


def find_equipment(query: str):
    """Cari equipment berdasarkan tag number dari query teks bebas."""
    if not os.path.exists(DB_PATH):
        return None

    conn   = get_connection()
    cursor = conn.cursor()

    try:
        query_clean   = query.upper().strip()
        query_nospace = query_clean.replace(" ", "").replace("-", "")

        # Tahap 1: exact match
        cursor.execute('''
            SELECT Tag_Number, PID_Number, Service_Description, PID_Link, Logic_Diagram, HMI_Link
            FROM equipment
            WHERE REPLACE(REPLACE(UPPER(Tag_Number), ' ', ''), '-', '') = ?
            LIMIT 1
        ''', (query_nospace,))
        row = cursor.fetchone()
        if row:
            return _to_dict(row)

        # Tahap 2: LIKE match
        cursor.execute('''
            SELECT Tag_Number, PID_Number, Service_Description, PID_Link, Logic_Diagram, HMI_Link
            FROM equipment
            WHERE REPLACE(REPLACE(UPPER(Tag_Number), ' ', ''), '-', '') LIKE ?
            LIMIT 1
        ''', (f'%{query_nospace}%',))
        row = cursor.fetchone()
        if row:
            return _to_dict(row)

        # Tahap 3: kombinasi token
        tokens = re.findall(r'[A-Z0-9]+', query_clean)
        for i in range(len(tokens)):
            for j in range(i+1, min(i+4, len(tokens)+1)):
                combo = "".join(tokens[i:j])
                if re.search(r'[A-Z]', combo) and re.search(r'\d', combo):
                    cursor.execute('''
                        SELECT Tag_Number, PID_Number, Service_Description, PID_Link, Logic_Diagram, HMI_Link
                        FROM equipment
                        WHERE REPLACE(REPLACE(UPPER(Tag_Number), ' ', ''), '-', '') = ?
                        LIMIT 1
                    ''', (combo,))
                    row = cursor.fetchone()
                    if row:
                        return _to_dict(row)

        # Tahap 4: fallback Service_Description
        keywords = [w for w in query_clean.split() if len(w) > 3]
        for kw in keywords:
            cursor.execute('''
                SELECT Tag_Number, PID_Number, Service_Description, PID_Link, Logic_Diagram, HMI_Link
                FROM equipment
                WHERE UPPER(Service_Description) LIKE ?
                LIMIT 1
            ''', (f'%{kw}%',))
            row = cursor.fetchone()
            if row:
                return _to_dict(row)

    except sqlite3.Error as e:
        print(f"Database error: {e}")
    finally:
        conn.close()

    return None


def get_all_equipment():
    """Ambil semua equipment sebagai list of dict."""
    if not os.path.exists(DB_PATH):
        return []
    conn   = get_connection()
    cursor = conn.cursor()
    cursor.execute('''
        SELECT Tag_Number, PID_Number, Service_Description, PID_Link, Logic_Diagram, HMI_Link
        FROM equipment ORDER BY Tag_Number
    ''')
    rows = cursor.fetchall()
    conn.close()
    return [_to_dict(r) for r in rows]


def get_all_tags():
    """Ambil semua tag number sebagai list string."""
    if not os.path.exists(DB_PATH):
        return []
    conn   = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT Tag_Number FROM equipment ORDER BY Tag_Number")
    tags = [r[0] for r in cursor.fetchall()]
    conn.close()
    return tags


# ─────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────
if __name__ == "__main__":
    args        = sys.argv[1:]
    reset       = "--reset"  in args
    verify_only = "--verify" in args

    print("=" * 50)
    print("AMMONIA PLANT — DATABASE SETUP")
    print("=" * 50)

    if verify_only:
        verify()
        sys.exit(0)

    create_database(reset=reset)
    imported = import_from_excel()

    if imported == 0:
        print("\nImport Excel gagal atau file tidak ditemukan.")
        print("Isi data manual via DB Browser for SQLite.")

    verify()

    print(f"\nSelesai! Database tersimpan di: {DB_PATH}")
    print("   Edit data       : DB Browser for SQLite")
    print("   Jalankan bot    : python main.py")
    print("   Jalankan backend: python backend.py")
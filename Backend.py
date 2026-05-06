"""
backend.py — FastAPI Backend untuk Ammonia Plant P&ID Search Website

Cara jalankan:
    pip install fastapi uvicorn
    python backend.py

Endpoint tersedia:
    GET  /equipment         → semua equipment
    GET  /search?q=ZC1016  → cari berdasarkan tag / service
    GET  /equipment/{tag}   → detail satu equipment
    GET  /health            → cek status server
"""

from fastapi import FastAPI, HTTPException, Query
from fastapi.middleware.cors import CORSMiddleware
import uvicorn
import sqlite3
import os
import re

DB_PATH = "data/ammonia.db"

def get_connection():
    return sqlite3.connect(DB_PATH, check_same_thread=False)

def _to_dict(row):
    return {
        'tag_number':           row[0],
        'pid_number':           row[1],
        'service_description':  row[2],
        'pid_link':             row[3],
        'Logic_Diagram':        row[4],
        'hmi_link':             row[5],
    }

# ── Kolom asli di DB ────────────────────────────────────────
# Tag_Number, PID_Number, Service_Description,
# PID_Link, Logic_Diagram, HMI_Link

SELECT_COLS = """
    Tag_Number,
    PID_Number,
    Service_Description,
    PID_Link,
    Logic_Diagram,
    HMI_Link
"""

def get_all_equipment():
    if not os.path.exists(DB_PATH):
        return []
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute(f'''
        SELECT {SELECT_COLS}
        FROM equipment
        ORDER BY Tag_Number
    ''')
    rows = cursor.fetchall()
    conn.close()
    return [_to_dict(r) for r in rows]

def get_all_tags():
    if not os.path.exists(DB_PATH):
        return []
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT Tag_Number FROM equipment ORDER BY Tag_Number")
    tags = [r[0] for r in cursor.fetchall()]
    conn.close()
    return tags

def find_equipment(query: str):
    if not os.path.exists(DB_PATH):
        return None
    conn   = get_connection()
    cursor = conn.cursor()
    try:
        query_clean   = query.upper().strip()
        query_nospace = query_clean.replace(" ", "").replace("-", "")

        # ── Exact match (setelah normalisasi) ──
        cursor.execute(f'''
            SELECT {SELECT_COLS}
            FROM equipment
            WHERE REPLACE(REPLACE(UPPER(Tag_Number), ' ', ''), '-', '') = ?
            LIMIT 1
        ''', (query_nospace,))
        row = cursor.fetchone()
        if row:
            return _to_dict(row)

        # ── Partial match tag_number ──
        cursor.execute(f'''
            SELECT {SELECT_COLS}
            FROM equipment
            WHERE REPLACE(REPLACE(UPPER(Tag_Number), ' ', ''), '-', '') LIKE ?
            LIMIT 1
        ''', (f'%{query_nospace}%',))
        row = cursor.fetchone()
        if row:
            return _to_dict(row)

        # ── Token combination match ──
        tokens = re.findall(r'[A-Z0-9]+', query_clean)
        for i in range(len(tokens)):
            for j in range(i + 1, min(i + 4, len(tokens) + 1)):
                combo = "".join(tokens[i:j])
                if re.search(r'[A-Z]', combo) and re.search(r'\d', combo):
                    cursor.execute(f'''
                        SELECT {SELECT_COLS}
                        FROM equipment
                        WHERE REPLACE(REPLACE(UPPER(Tag_Number), ' ', ''), '-', '') = ?
                        LIMIT 1
                    ''', (combo,))
                    row = cursor.fetchone()
                    if row:
                        return _to_dict(row)

        # ── Keyword match di service_description ──
        keywords = [w for w in query_clean.split() if len(w) > 3]
        for kw in keywords:
            cursor.execute(f'''
                SELECT {SELECT_COLS}
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

# ─────────────────────────────────────────────
# INISIALISASI APP
# ─────────────────────────────────────────────
app = FastAPI(
    title="Ammonia Plant P&ID Search API",
    description="Backend API untuk sistem pencarian P&ID Ammonia Plant",
    version="1.0.0"
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["GET"],
    allow_headers=["*"],
)

# ─────────────────────────────────────────────
# ENDPOINTS
# ─────────────────────────────────────────────

@app.get("/health")
def health_check():
    """Cek apakah server berjalan."""
    tags = get_all_tags()
    return {
        "status": "ok",
        "total_equipment": len(tags),
        "message": "Ammonia Plant API berjalan normal"
    }


@app.get("/equipment")
def get_all():
    """Ambil semua equipment dari database."""
    data = get_all_equipment()
    if not data:
        return []
    return data


@app.get("/search")
def search(q: str = Query(..., min_length=1, description="Tag number atau keyword")):
    q_upper   = q.upper().strip()
    q_nospace = q_upper.replace(" ", "").replace("-", "")

    conn   = get_connection()
    cursor = conn.cursor()

    try:
        cursor.execute(f'''
            SELECT {SELECT_COLS}
            FROM equipment
            WHERE REPLACE(REPLACE(UPPER(Tag_Number), ' ', ''), '-', '') LIKE ?
               OR UPPER(Service_Description) LIKE ?
               OR UPPER(PID_Number) LIKE ?
            ORDER BY Tag_Number
            LIMIT 100
        ''', (
            f'%{q_nospace}%',
            f'%{q_upper}%',
            f'%{q_upper}%',
        ))
        rows = cursor.fetchall()
        return [_to_dict(r) for r in rows]

    except sqlite3.Error as e:
        raise HTTPException(status_code=500, detail=f"Database error: {e}")
    finally:
        conn.close()


@app.get("/equipment/{tag}")
def get_by_tag(tag: str):
    """Ambil detail satu equipment berdasarkan tag number."""
    result = find_equipment(tag)
    if not result:
        raise HTTPException(status_code=404, detail=f"Tag '{tag}' tidak ditemukan")
    return result


# ─────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────
if __name__ == "__main__":
    print("=" * 50)
    print("AMMONIA PLANT — BACKEND API")
    print("=" * 50)
    print("Server berjalan di: http://localhost:8000")
    print("Dokumentasi API   : http://localhost:8000/docs")
    print("Tekan Ctrl+C untuk stop.")
    print("=" * 50)

    uvicorn.run(app, host="0.0.0.0", port=8000)
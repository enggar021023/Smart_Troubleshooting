# Smart Troubleshooting — Ammonia Plant P&ID Search

Sistem pencarian P&ID (Piping & Instrumentation Diagram) dan informasi equipment untuk Ammonia Plant PT Petrokimia Gresik.

## 🎯 Fitur

- **Website P&ID Search** — Interface web modern untuk mencari equipment
- **FastAPI Backend** — API REST untuk data equipment dan P&ID
- **Telegram Bot** — Bot Telegram untuk query cepat via messaging
- **Database SQLite** — Manajemen data equipment lokal
- **Dark/Light Mode** — Theme toggle untuk kenyamanan user

## 📁 Struktur Project

```
Smart_Troubleshooting/
├── Backend.py              # FastAPI backend (port 8000)
├── main.py                 # Telegram bot
├── frontend/
│   └── index.html          # Website P&ID search
├── data/
│   └── ammonia.db          # SQLite database
├── setup_db.py             # Script setup database
├── requirements.txt        # Python dependencies
├── docker-compose.yml      # Docker orchestration
└── README.md               # File ini
```

## 🚀 Quick Start

### 1. Install Dependencies
```bash
pip install -r requirements.txt
```

### 2. Setup Database (jalankan sekali)
```bash
python setup_db.py
```

### 3. Jalankan Backend
```bash
python Backend.py
```

Server akan running di: `http://localhost:8000`

### 4. Akses Website
Buka file `frontend/index.html` di browser, atau serve via HTTP:
```bash
# Via Python
python -m http.server 8001 --directory frontend

# Atau buka langsung: file:///path/to/frontend/index.html
```

### 5. Jalankan Telegram Bot (opsional)
```bash
python main.py
```

Perlu set environment variables di `.env`:
```
TELEGRAM_TOKEN=your_token_here
GROQ_API_KEY=your_api_key_here
```

## 📡 API Endpoints

| Method | Endpoint | Deskripsi |
|--------|----------|-----------|
| GET | `/health` | Status server & total equipment |
| GET | `/equipment` | Semua equipment |
| GET | `/search?q=ZC1016` | Cari tag/service |
| GET | `/equipment/{tag}` | Detail equipment |

## 🔧 Teknologi

- **Backend**: FastAPI, Uvicorn
- **Frontend**: HTML5, CSS3, JavaScript (Vanilla)
- **Database**: SQLite3
- **Bot**: python-telegram-bot
- **AI**: Groq (fallback untuk queries tidak ditemukan)

## 📝 Environment Variables

Buat file `.env`:
```
TELEGRAM_TOKEN=your_telegram_bot_token
GROQ_API_KEY=your_groq_api_key
```

## 🐳 Docker

Jalankan seluruh stack via Docker Compose:
```bash
docker-compose up
```

## 📊 Database

Edit database via:
- **DB Browser for SQLite** (GUI tool)
- Atau direct SQL commands

Struktur tabel `equipment`:
```sql
CREATE TABLE equipment (
    id INTEGER PRIMARY KEY,
    Tag_Number TEXT,
    PID_Number TEXT,
    Service_Description TEXT,
    PID_Link TEXT,
    Logic_Diagram TEXT,
    HMI_Link TEXT
);
```

## 🎨 Features

- ✅ Search real-time dengan highlight keyword
- ✅ Filter by instrument type (Pressure, Temperature, Flow, dll)
- ✅ Filter by documentation completeness
- ✅ Dark/Light theme toggle
- ✅ Search history
- ✅ Copy tag to clipboard
- ✅ Responsive design (mobile-friendly)

## 📧 Support

Untuk pertanyaan atau bug report, hubungi team development.

---

**PT Petrokimia Gresik — Ammonia Plant**
**v1.0.0**

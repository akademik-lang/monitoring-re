"""
generate_excel.py
=================
Generator Excel monitoring mahasiswa TOEFL dari file ODS.

Fitur:
  - Baca semua kelas dari data.json
  - Nilai tiap session (Grammar, Listening, Reading, Test Prep) rapi
  - Threshold TOEFL >= 403 → layak sertifikat
  - Cek otomatis file sertifikat di Google Drive per kelas
  - Cek nama mahasiswa ke SIAKAD (sip.sains.ac.id) by NIM
  - Cache SIAKAD: nama yang sudah "Cocok" di-skip, yang bermasalah di-re-cek
  - 1 file Excel, multiple sheet: RINGKASAN + 1 sheet per kelas
  - Penamaan sheet unik dari path folder

Cara pakai:
  python generate_excel.py          # Normal: skip nama yang sudah "Cocok"
  python generate_excel.py --full   # Cek semua mahasiswa dari awal (abaikan cache)

Credential: auth/credentials.json (OAuth2 installed)
            auth/siakad_session.json (cookie SIAKAD)
            auth/siakad_cache.json (cache hasil SIAKAD, auto-generated)
"""

import json
import os
import re
import sys
import time
from datetime import datetime

import requests
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── Google API ─────────────────────────────────────────────────────────────────
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

# ─── KONFIGURASI ──────────────────────────────────────────────────────────────
BASE_DIR      = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR    = os.path.join(BASE_DIR, "output")
DATA_JSON     = os.path.join(BASE_DIR, "data.json")
CRED_FILE     = os.path.join(BASE_DIR, "auth", "credentials.json")
TOKEN_FILE    = os.path.join(BASE_DIR, "auth", "token.json")
TOEFL_PASS    = 403   # Nilai minimum layak sertifikat

# SIAKAD
SIAKAD_CONFIG_FILE = os.path.join(BASE_DIR, "auth", "siakad_session.json")
SIAKAD_URL   = "https://sip.sains.ac.id/mahasiswa/listdata"
SIAKAD_CACHE_FILE = os.path.join(BASE_DIR, "auth", "siakad_cache.json")

# Scope: baca metadata file Drive (tidak perlu download)
SCOPES = ["https://www.googleapis.com/auth/drive.readonly"]

os.makedirs(OUTPUT_DIR, exist_ok=True)

# ─── WARNA ────────────────────────────────────────────────────────────────────
C = {
    "TITLE_BG"    : "0D1B2A",
    "TITLE_FG"    : "F1C40F",
    "HEADER_BG"   : "1E3A5F",
    "HEADER_FG"   : "FFFFFF",
    "NIM_BG"      : "2E86AB",
    "NIM_FG"      : "FFFFFF",
    "DRIVE_BG"    : "6C3483",
    "DRIVE_FG"    : "FFFFFF",
    "ROW_ODD"     : "F0F4F8",
    "ROW_EVEN"    : "FFFFFF",
    "SCORE_LOW"   : "FFCCCC",   # < 60
    "SCORE_MED"   : "FFF3CC",   # 60-79
    "SCORE_HIGH"  : "D4EDDA",   # >= 80
    "CERT_YES"    : "27AE60",   # Hijau – ada sertifikat
    "CERT_NO"     : "E74C3C",   # Merah – belum ada
    "CERT_SKIP"   : "BDC3C7",   # Abu – tidak layak
    "PASS_BG"     : "D5F5E3",   # Hijau muda – lulus
    "FAIL_BG"     : "FADBD8",   # Merah muda – remedial
    "SUMMARY_BG"  : "2C3E50",
    "SUMMARY_FG"  : "F1C40F",
    "SESSION_BG"  : "EBF5FB",   # Biru sangat muda – kolom sesi
}

thin_border = Border(
    left=Side(style='thin', color='CCCCCC'),
    right=Side(style='thin', color='CCCCCC'),
    top=Side(style='thin', color='CCCCCC'),
    bottom=Side(style='thin', color='CCCCCC'),
)

def fill(hex_color):
    return PatternFill(fill_type="solid", fgColor=hex_color)

def fnt(color="000000", bold=False, size=10, underline=None):
    return Font(color=color, bold=bold, size=size, name="Calibri",
                underline=underline or 'none')

def hdr(cell, bg=None, fg=None, bold=True, size=10, wrap=True):
    cell.fill = fill(bg or C["HEADER_BG"])
    cell.font = fnt(fg or C["HEADER_FG"], bold=bold, size=size)
    cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=wrap)
    cell.border = thin_border

def dat(cell, bg=None, fg="000000", bold=False, size=10, align='center', wrap=False):
    if bg:
        cell.fill = fill(bg)
    cell.font = fnt(fg, bold=bold, size=size)
    cell.alignment = Alignment(horizontal=align, vertical='center', wrap_text=wrap)
    cell.border = thin_border

def score_color(val):
    try:
        v = float(str(val).replace(',', '.'))
        if v < 60:   return C["SCORE_LOW"]
        elif v < 80: return C["SCORE_MED"]
        else:        return C["SCORE_HIGH"]
    except:
        return None

def safe(v):
    s = str(v).strip()
    return '-' if s in ['nan', 'NaT', 'None', ''] else s

def get_name(row):
    first   = str(row.get('First name', '')).strip()
    surname = str(row.get('Surname', '')).strip()
    if surname in ['.', '-', 'nan', '']:
        surname = ''
    surname = re.sub(r'\[NIM\s*:.*?\]', '', surname).strip()
    parts = [p for p in [first, surname] if p]
    return ' '.join(parts)

def make_sheet_name(file_path):
    norm  = file_path.replace('\\', '/').replace('./', '')
    parts = norm.split('/')
    try:
        idx      = parts.index('resource')
        relevant = parts[idx + 1:]
    except ValueError:
        relevant = parts[-3:]
    name_parts = [p.upper().replace('.ODS', '') for p in relevant]
    return '_'.join(name_parts)[:31]


# ─── GOOGLE DRIVE AUTH ────────────────────────────────────────────────────────
def get_drive_service():
    """Autentikasi ke Google Drive dan kembalikan service object."""
    creds = None

    if os.path.exists(TOKEN_FILE):
        creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            print("  Refreshing token...")
            creds.refresh(Request())
        else:
            print("  Membuka browser untuk otorisasi Google Drive...")
            flow = InstalledAppFlow.from_client_secrets_file(CRED_FILE, SCOPES)
            creds = flow.run_local_server(port=0)

        with open(TOKEN_FILE, 'w') as token:
            token.write(creds.to_json())
        print("  Token disimpan.")

    service = build('drive', 'v3', credentials=creds)
    return service


def extract_folder_id(drive_url):
    """Ambil folder ID dari Google Drive URL."""
    # Format: https://drive.google.com/drive/folders/{ID}?...
    match = re.search(r'/folders/([a-zA-Z0-9_-]+)', drive_url)
    if match:
        return match.group(1)
    # Format: https://drive.google.com/open?id={ID}
    match = re.search(r'[?&]id=([a-zA-Z0-9_-]+)', drive_url)
    if match:
        return match.group(1)
    return None


def list_drive_files(service, folder_id):
    """List semua file dalam folder Drive. Return set nama file (lowercase)."""
    if not folder_id:
        return set()
    try:
        results = service.files().list(
            q=f"'{folder_id}' in parents and trashed=false",
            fields="files(id, name, mimeType)",
            pageSize=1000
        ).execute()
        files = results.get('files', [])
        # Return set nama file (lowercase, tanpa ekstensi untuk fleksibilitas matching)
        return {f['name'] for f in files}, files
    except Exception as e:
        print(f"  WARNING: Gagal list Drive folder {folder_id}: {e}")
        return set(), []


def is_cert_in_drive(student_name, drive_files_set):
    """
    Cek apakah sertifikat mahasiswa ada di Drive.
    Matching by nama mahasiswa (case-insensitive, partial match).
    """
    if not drive_files_set:
        return None  # Tidak bisa dicek (error Drive)

    name_lower = student_name.lower().strip()
    # Ambil kata-kata dari nama
    name_words = set(name_lower.split())

    for fname in drive_files_set:
        fname_lower = fname.lower()
        # Cek apakah semua kata nama ada di nama file
        if all(w in fname_lower for w in name_words if len(w) > 2):
            return True
        # Atau exact match setelah normalisasi
        fname_clean = re.sub(r'[^a-z0-9\s]', '', fname_lower)
        name_clean  = re.sub(r'[^a-z0-9\s]', '', name_lower)
        if name_clean and name_clean in fname_clean:
            return True

    return False


# ─── SIAKAD API ───────────────────────────────────────────────────────────────
_siakad_cache = {}   # cache NIM → nama SIAKAD (hindari request berulang)

def load_siakad_config():
    """Load cookie session SIAKAD dari auth/siakad_session.json."""
    if not os.path.exists(SIAKAD_CONFIG_FILE):
        return None
    try:
        with open(SIAKAD_CONFIG_FILE, 'r') as f:
            cfg = json.load(f)
        if not cfg.get('csrf_cookie_name') or not cfg.get('siakad_session'):
            print("    WARNING: Cookie SIAKAD kosong. Update auth/siakad_session.json.")
            return None
        return cfg
    except Exception as e:
        print(f"    WARNING: Gagal load config SIAKAD: {e}")
        return None


def load_siakad_cache():
    """Load persistent cache hasil SIAKAD dari file."""
    if not os.path.exists(SIAKAD_CACHE_FILE):
        return {}
    try:
        with open(SIAKAD_CACHE_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception:
        return {}


def save_siakad_cache(cache):
    """Simpan cache hasil SIAKAD ke file."""
    try:
        with open(SIAKAD_CACHE_FILE, 'w', encoding='utf-8') as f:
            json.dump(cache, f, indent=2, ensure_ascii=False)
    except Exception as e:
        print(f"    WARNING: Gagal simpan cache SIAKAD: {e}")


def build_siakad_body(nim, draw=4):
    """Bangun body request DataTables untuk search by NIM."""
    body = {}
    # 9 kolom (index 0-8), kolom 0 tidak orderable
    for i in range(9):
        body[f"columns[{i}][data]"]      = str(i)
        body[f"columns[{i}][name]"]      = ""
        body[f"columns[{i}][searchable]"] = "true"
        body[f"columns[{i}][orderable]"]  = "false" if i == 0 else "true"
        body[f"columns[{i}][search][value]"] = ""
        body[f"columns[{i}][search][regex]"] = "false"
    body["draw"]            = str(draw)
    body["order[0][column]"] = "3"
    body["order[0][dir]"]    = "asc"
    body["start"]            = "0"
    body["length"]           = "10"
    body["search[value]"]    = str(nim)
    body["search[regex]"]    = "false"
    body["jenjangprodi"]     = ""
    body["lokasi"]           = ""
    body["angkatan"]         = ""
    body["status"]           = ""
    body["jk"]               = ""
    body["agama"]            = ""
    body["id_kls"]           = ""
    return body


def fetch_siakad_name(nim, siakad_config):
    """
    Query SIAKAD untuk cari mahasiswa by NIM.
    Return: nama mahasiswa (str) atau None jika tidak ketemu / error.
    """
    if not siakad_config or not nim:
        return None

    # Cek cache
    if nim in _siakad_cache:
        return _siakad_cache[nim]

    cookies = {
        "csrf_cookie_name": siakad_config.get("csrf_cookie_name", ""),
        "siakad_session":   siakad_config.get("siakad_session", ""),
    }
    headers = {
        "Accept":          "application/json, text/javascript, */*; q=0.01",
        "Content-Type":     "application/x-www-form-urlencoded; charset=UTF-8",
        "X-Requested-With": "XMLHttpRequest",
        "User-Agent":       "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                            "AppleWebKit/537.36 (KHTML, like Gecko) "
                            "Chrome/156.0.0.0 Safari/537.36",
        "Origin":   "https://sip.sains.ac.id",
        "Referer":  "https://sip.sains.ac.id/mahasiswa",
    }

    try:
        resp = requests.post(
            SIAKAD_URL,
            cookies=cookies,
            headers=headers,
            data=build_siakad_body(nim),
            timeout=15,
        )
        resp.raise_for_status()

        # Cek apakah response JSON (bukan redirect ke login)
        try:
            result = resp.json()
        except ValueError:
            print(f"      SIAKAD: Response bukan JSON (mungkin session expired). NIM={nim}")
            _siakad_cache[nim] = None
            return None

        rows = result.get("data", [])
        if not rows:
            _siakad_cache[nim] = None
            return None

        # Cari baris yang NIM-nya cocok (index 2 = NIM, index 3 = nama)
        for row in rows:
            if len(row) >= 4 and str(row[2]).strip() == str(nim).strip():
                name = str(row[3]).strip()
                _siakad_cache[nim] = name
                return name

        # Fallback: ambil baris pertama
        name = str(rows[0][3]).strip()
        _siakad_cache[nim] = name
        return name

    except requests.exceptions.RequestException as e:
        print(f"      SIAKAD ERROR NIM={nim}: {e}")
        _siakad_cache[nim] = None
        return None


def normalize_name(name):
    """Normalisasi nama untuk perbandingan."""
    if not name:
        return ""
    n = name.strip().lower()
    n = re.sub(r'\s+', ' ', n)               # collapse multiple spaces
    # Hapus gelar di depan
    n = re.sub(
        r'^(dr\.|prof\.|ir\.|mr\.|mrs\.|s\.t\.|s\.kom\.|s\.si\.|m\.t\.|m\.kom\.|s\.pd\.|s\.pd)\s*',
        '', n
    )
    return n.strip()


def compare_names(name_ods, name_siakad):
    """
    Bandingkan nama dari ODS dengan nama dari SIAKAD.
    Return: (match: bool|None, detail: str)
      True  → cocok
      False → beda
      None  → tidak bisa cek (SIAKAD error / tidak ketemu)
    """
    if not name_siakad:
        return None, "N/A"

    n1 = normalize_name(name_ods)
    n2 = normalize_name(name_siakad)

    if not n1 or not n2:
        return None, "N/A"

    # Exact match setelah normalisasi
    if n1 == n2:
        return True, "Cocok"

    # Cek contains (salah satu mengandung yang lain)
    if n1 in n2 or n2 in n1:
        return True, "Cocok"

    # Cek kata-per-kata (urutan beda tapi kata sama)
    words1 = set(n1.split())
    words2 = set(n2.split())
    if words1 == words2:
        return True, "Cocok"

    return False, "Beda"


# ─── PARSE SCORM COLUMNS ──────────────────────────────────────────────────────
def parse_scorm_columns(headers):
    """
    Kelompokkan kolom SCORM per session.
    Return: dict { session_num: { sub_type: col_name } }
    Sub types: Grammar, Listening, Reading, Test Prep, Study Tips, Pretest, Final Test, Susulan, Preparation
    """
    session_pattern = re.compile(
        r'SCORM package:\s*Session\s*(\d+)\s*-\s*([^-]+?)\s*-\s*(.+?)\s*\(Real\)',
        re.IGNORECASE
    )
    special_pattern = re.compile(
        r'SCORM package:\s*(STUDY TIPS|PRETEST|FINAL TEST|SUSULAN.*?FINAL TEST|PREPARATION.*?)\s*-\s*(.+?)\s*\(Real\)',
        re.IGNORECASE
    )

    sessions = {}      # { int: { label: col_name } }
    specials = {}      # { label: col_name }

    for h in headers:
        h_str = str(h)
        m = session_pattern.match(h_str)
        if m:
            num  = int(m.group(1))
            sub  = m.group(2).strip()   # e.g. "Grammar", "Listening", etc.
            # Normalize sub-type
            sub_norm = normalize_subtype(sub)
            if num not in sessions:
                sessions[num] = {}
            sessions[num][sub_norm] = h_str
            continue

        m2 = special_pattern.match(h_str)
        if m2:
            label = m2.group(1).strip().upper()
            specials[label] = h_str

    return sessions, specials


def normalize_subtype(sub):
    """Normalize sub-type name."""
    s = sub.lower().strip()
    if 'grammar'      in s: return 'Grammar'
    if 'listening'    in s: return 'Listening'
    if 'reading'      in s: return 'Reading'
    if 'test prep'    in s: return 'Test Prep'
    if 'preparation'  in s: return 'Test Prep'
    return sub.title()


# ─── BUILD SHEETS ─────────────────────────────────────────────────────────────
def read_ods(file_path):
    df  = pd.read_excel(file_path, engine='odf', sheet_name=None, header=None)
    raw = list(df.values())[0]
    headers = raw.iloc[0].tolist()
    data    = raw.iloc[1:].copy()
    data.columns = headers
    data    = data.reset_index(drop=True)
    return data, headers


def build_class_sheet(wb, file_info, drive_service, siakad_config=None, siakad_cache=None, full_check=False):
    """Buat sheet detail satu kelas. Return dict info untuk ringkasan."""
    file_path  = os.path.join(BASE_DIR, file_info['filePath'].replace('./', ''))
    drive_url  = file_info['googleDriveUrl']
    sheet_name = make_sheet_name(file_info['filePath'])
    folder_id  = extract_folder_id(drive_url)

    # Info kelas dari path
    norm  = file_info['filePath'].replace('\\', '/').replace('./', '')
    parts = norm.split('/')
    try:
        idx      = parts.index('resource')
        tahun    = parts[idx + 1]
        semester = parts[idx + 2].upper()
        kelas    = os.path.splitext(parts[-1])[0].upper()
    except:
        tahun = semester = ""; kelas = sheet_name

    print(f"  [{kelas}] Membaca ODS...")
    data, headers = read_ods(file_path)

    print(f"  [{kelas}] Listing Google Drive folder...")
    drive_files_set, drive_files_raw = list_drive_files(drive_service, folder_id)
    print(f"  [{kelas}] File di Drive: {len(drive_files_set)}")
    if drive_files_raw:
        print(f"           Contoh: {[f['name'] for f in drive_files_raw[:5]]}")

    sessions, specials = parse_scorm_columns(headers)
    session_nums = sorted(sessions.keys())

    # ── Buat worksheet ────────────────────────────────────────────────────────
    ws = wb.create_sheet(sheet_name)

    # Hitung total kolom
    # Kolom awal: No | Nama | NIM | Email
    # Per session: Grammar | Listening | Reading | Test Prep
    # Kolom akhir: TOEFL | Layak Sertifikat? | Status Sertifikat di Drive | Keterangan
    SUBTYPES = ['Grammar', 'Listening', 'Reading', 'Test Prep']
    fixed_start = 4   # No, Nama, NIM, Email
    scorm_count = len(session_nums) * len(SUBTYPES)
    fixed_end   = 6   # TOEFL, Layak?, Drive, Nama SIAKAD, Nama Cocok?, Keterangan
    total_cols  = fixed_start + scorm_count + fixed_end
    last_col    = get_column_letter(total_cols)

    # ── Baris 1: Judul ────────────────────────────────────────────────────────
    ws.merge_cells(f'A1:{last_col}1')
    c = ws['A1']
    c.value = f"MONITORING MAHASISWA  |  KELAS {kelas}  |  {semester} {tahun}"
    c.fill = fill(C["TITLE_BG"])
    c.font = Font(color=C["TITLE_FG"], bold=True, size=14, name="Calibri")
    c.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 38

    # ── Baris 2: Drive URL ────────────────────────────────────────────────────
    ws.merge_cells(f'A2:{last_col}2')
    c2 = ws['A2']
    c2.value = f"Google Drive Kelas {kelas}  →  {drive_url}"
    c2.fill = fill(C["DRIVE_BG"])
    c2.font = Font(color=C["DRIVE_FG"], bold=False, size=10,
                    underline='single', name="Calibri")
    c2.alignment = Alignment(horizontal='center', vertical='center')
    c2.hyperlink = drive_url
    ws.row_dimensions[2].height = 22

    # ── Baris 3-4: Header 2 level ─────────────────────────────────────────────
    # Baris 3: Grup (No, Nama, NIM, Email | Session 1 ... Session N | Akhir)
    # Baris 4: Sub-header

    # Kolom fixed awal
    fixed_start_labels = ["No", "Nama Mahasiswa", "NIM", "Email"]
    for ci, label in enumerate(fixed_start_labels, 1):
        ws.merge_cells(start_row=3, start_column=ci, end_row=4, end_column=ci)
        cell = ws.cell(row=3, column=ci)
        cell.value = label
        hdr(cell, size=10)
        ws.cell(row=4, column=ci).border = thin_border

    # Kolom session
    session_col_start = fixed_start + 1
    for si, snum in enumerate(session_nums):
        col_start = session_col_start + si * len(SUBTYPES)
        col_end   = col_start + len(SUBTYPES) - 1

        # Merge row 3 untuk judul session
        ws.merge_cells(start_row=3, start_column=col_start, end_row=3, end_column=col_end)
        hdr_cell = ws.cell(row=3, column=col_start)
        hdr_cell.value = f"Session {snum}"
        hdr(hdr_cell, bg="154360", size=9)  # Biru tua untuk session header

        # Sub-header row 4
        for bi, sub in enumerate(SUBTYPES):
            sub_cell = ws.cell(row=4, column=col_start + bi)
            sub_cell.value = sub
            hdr(sub_cell, bg="1A5276", size=9)

    # Kolom fixed akhir
    end_col_start = session_col_start + len(session_nums) * len(SUBTYPES)
    end_labels = [
        "TOEFL\nScore",
        f"Layak\nSertifikat?\n(≥{TOEFL_PASS})",
        "Sertifikat\ndi Drive",
        "Nama di\nSIAKAD",
        "Nama\nCocok?",
        "Keterangan"
    ]
    end_bgs = [C["HEADER_BG"], "1E8449", "6C3483", "1A5276", "1A5276", C["HEADER_BG"]]
    for bi, (label, bg) in enumerate(zip(end_labels, end_bgs)):
        ci = end_col_start + bi
        ws.merge_cells(start_row=3, start_column=ci, end_row=4, end_column=ci)
        cell = ws.cell(row=3, column=ci)
        cell.value = label
        hdr(cell, bg=bg, size=10)
        ws.cell(row=4, column=ci).border = thin_border

    ws.row_dimensions[3].height = 38
    ws.row_dimensions[4].height = 30

    # ── Data ──────────────────────────────────────────────────────────────────
    total_students = len(data)
    layak_count    = 0
    cert_yes_count = 0
    cert_no_count  = 0
    cert_unk_count = 0
    name_match_count    = 0
    name_mismatch_count = 0
    name_unk_count      = 0

    for ri, (_, row) in enumerate(data.iterrows(), 1):
        er = 4 + ri
        bg = C["ROW_ODD"] if ri % 2 == 0 else C["ROW_EVEN"]

        full_name = get_name(row)
        nim       = str(row.get('Username', '')).strip()
        email     = str(row.get('Email address', '')).strip()

        # ── Cek nama ke SIAKAD ────────────────────────────────────────────────
        siakad_name = None
        name_match = None
        if siakad_config:
            cached = siakad_cache.get(nim) if siakad_cache and not full_check else None

            if cached and cached.get("last_match") is True and cached.get("siakad_name"):
                # Sudah cocok sebelumnya → skip API call, pakai cache
                siakad_name = cached["siakad_name"]
                name_match, _ = compare_names(full_name, siakad_name)
                tag = "✅" if name_match else "❌"
                print(f"      [{ri}/{total_students}] NIM {nim} → cache: {siakad_name} ({tag})")
            else:
                # Belum di-cache / bermasalah / full check → query SIAKAD
                label = "re-cek" if cached else "cek"
                print(f"      [{ri}/{total_students}] NIM {nim} → {label} SIAKAD...", end='', flush=True)
                siakad_name = fetch_siakad_name(nim, siakad_config)
                name_match, _ = compare_names(full_name, siakad_name)
                if siakad_name:
                    status = "✅ cocok" if name_match else "❌ BEDA"
                    print(f" {siakad_name} ({status})")
                else:
                    print(" tidak ditemukan / error")
                time.sleep(0.3)  # hindari rate-limit

            # Update cache
            if siakad_cache is not None:
                siakad_cache[nim] = {"siakad_name": siakad_name, "last_match": name_match}
        else:
            name_match = None

        if name_match is True:    name_match_count += 1
        elif name_match is False: name_mismatch_count += 1
        else:                     name_unk_count += 1

        # TOEFL
        toefl_raw = safe(row.get('TOEFL (Real)', '-'))
        try:
            toefl_num = float(str(toefl_raw).replace(',', '.'))
            toefl_val = int(toefl_num) if toefl_num == int(toefl_num) else toefl_num
            layak     = toefl_num >= TOEFL_PASS
        except:
            toefl_val = toefl_raw
            layak     = False

        if layak:
            layak_count += 1

        # Cek Drive
        if layak and drive_files_set is not None:
            cert_found = is_cert_in_drive(full_name, drive_files_set)
        elif not layak:
            cert_found = None  # Tidak layak, skip cek
        else:
            cert_found = None  # Error Drive

        if cert_found is True:    cert_yes_count += 1
        elif cert_found is False: cert_no_count  += 1
        else:                     cert_unk_count += 1

        # ── Isi kolom fixed awal ──────────────────────────────────────────────
        fixed_vals = [ri, full_name, nim, email]
        for ci, val in enumerate(fixed_vals, 1):
            cell = ws.cell(row=er, column=ci)
            cell.value = val
            if ci == 1:
                dat(cell, bg=bg)
            elif ci == 2:
                dat(cell, bg=bg, align='left', bold=True)
            elif ci == 3:
                dat(cell, bg=C["NIM_BG"], fg=C["NIM_FG"], bold=True)
            else:
                dat(cell, bg=bg, align='left', fg="4A4A4A", size=9)

        # ── Isi kolom session ─────────────────────────────────────────────────
        for si, snum in enumerate(session_nums):
            col_start = session_col_start + si * len(SUBTYPES)
            for bi, sub in enumerate(SUBTYPES):
                ci   = col_start + bi
                cell = ws.cell(row=er, column=ci)
                col_name = sessions[snum].get(sub)
                val  = safe(row.get(col_name, '-')) if col_name else '-'
                cell.value = val if val != '-' else '-'
                sc = score_color(val)
                dat(cell, bg=sc if sc else C["SESSION_BG"], size=9)

        # ── Isi kolom fixed akhir ─────────────────────────────────────────────
        # TOEFL Score
        ci_toefl = end_col_start
        cell = ws.cell(row=er, column=ci_toefl)
        cell.value = toefl_val
        sc = score_color(toefl_val) if str(toefl_val) != '-' else None
        dat(cell, bg=sc if sc else bg, bold=True, size=10)

        # Layak Sertifikat?
        ci_layak = end_col_start + 1
        cell = ws.cell(row=er, column=ci_layak)
        if str(toefl_val) == '-':
            cell.value = "-"
            dat(cell, bg=C["CERT_SKIP"], fg="7F8C8D")
        elif layak:
            cell.value = f"✅ LAYAK"
            dat(cell, bg=C["PASS_BG"], fg="1E8449", bold=True)
        else:
            cell.value = f"❌ TIDAK"
            dat(cell, bg=C["FAIL_BG"], fg=C["CERT_NO"], bold=True)

        # Status Sertifikat di Drive
        ci_cert = end_col_start + 2
        cell = ws.cell(row=er, column=ci_cert)
        if not layak:
            cell.value = "—"
            dat(cell, bg=C["CERT_SKIP"], fg="7F8C8D")
        elif cert_found is True:
            cell.value = "✅ Ada"
            dat(cell, bg=C["CERT_YES"], fg="FFFFFF", bold=True)
        elif cert_found is False:
            cell.value = "❌ Belum"
            dat(cell, bg=C["CERT_NO"], fg="FFFFFF", bold=True)
        else:
            cell.value = "⚠️ N/A"
            dat(cell, bg=C["CERT_SKIP"], fg="5D6D7E")

        # Nama di SIAKAD
        ci_siakad = end_col_start + 3
        cell = ws.cell(row=er, column=ci_siakad)
        cell.value = siakad_name if siakad_name else "—"
        dat(cell, bg=bg, align='left', size=9,
            fg="000000" if siakad_name else "7F8C8D")

        # Nama Cocok?
        ci_match = end_col_start + 4
        cell = ws.cell(row=er, column=ci_match)
        if name_match is True:
            cell.value = "✅ Cocok"
            dat(cell, bg=C["PASS_BG"], fg="1E8449", bold=True)
        elif name_match is False:
            cell.value = "❌ Beda"
            dat(cell, bg=C["FAIL_BG"], fg=C["CERT_NO"], bold=True)
        else:
            cell.value = "⚠️ N/A"
            dat(cell, bg=C["CERT_SKIP"], fg="5D6D7E")

        # Keterangan
        ci_ket = end_col_start + 5
        cell = ws.cell(row=er, column=ci_ket)
        cell.value = ""
        dat(cell, bg=bg, align='left')

        ws.row_dimensions[er].height = 20

    # ── Summary baris bawah ───────────────────────────────────────────────────
    sum_row = 4 + total_students + 2
    ws.merge_cells(f'A{sum_row}:{last_col}{sum_row}')
    sc = ws[f'A{sum_row}']
    sc.value = (
        f"Total: {total_students} mhs  |  Layak Sertifikat: {layak_count}  |  "
        f"Sertifikat Sudah di Drive: {cert_yes_count}  |  "
        f"Belum di Drive: {cert_no_count}  |  "
        f"Nama Cocok: {name_match_count}  |  Nama Beda: {name_mismatch_count}  |  "
        f"Kelas: {kelas}"
    )
    sc.fill = fill(C["SUMMARY_BG"])
    sc.font = Font(color=C["SUMMARY_FG"], bold=True, size=10, name="Calibri")
    sc.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[sum_row].height = 26

    # ── Lebar kolom ───────────────────────────────────────────────────────────
    ws.column_dimensions['A'].width = 5
    ws.column_dimensions['B'].width = 26
    ws.column_dimensions['C'].width = 14
    ws.column_dimensions['D'].width = 28
    # Session columns
    for ci in range(5, 5 + scorm_count):
        ws.column_dimensions[get_column_letter(ci)].width = 11
    # End columns
    ws.column_dimensions[get_column_letter(end_col_start)].width = 10      # TOEFL
    ws.column_dimensions[get_column_letter(end_col_start + 1)].width = 14  # Layak
    ws.column_dimensions[get_column_letter(end_col_start + 2)].width = 14  # Drive
    ws.column_dimensions[get_column_letter(end_col_start + 3)].width = 24  # Nama SIAKAD
    ws.column_dimensions[get_column_letter(end_col_start + 4)].width = 12  # Nama Cocok
    ws.column_dimensions[get_column_letter(end_col_start + 5)].width = 22  # Keterangan

    ws.freeze_panes = ws.cell(row=5, column=5)

    return {
        'class_name'   : kelas,
        'semester'     : f"{semester} {tahun}",
        'total'        : total_students,
        'layak'        : layak_count,
        'cert_yes'     : cert_yes_count,
        'cert_no'      : cert_no_count,
        'drive_url'    : drive_url,
        'sheet_name'   : sheet_name,
        'drive_count'  : len(drive_files_set),
        'name_match'    : name_match_count,
        'name_mismatch' : name_mismatch_count,
    }


def build_summary_sheet(wb, all_class_data):
    """Sheet RINGKASAN – tampilan pertama."""
    ws = wb.active
    ws.title = "RINGKASAN"

    # Judul
    ws.merge_cells('A1:L1')
    c = ws['A1']
    c.value = "📊  MONITORING SERTIFIKAT TOEFL  |  SEMUA KELAS"
    c.fill = fill(C["TITLE_BG"])
    c.font = Font(color=C["TITLE_FG"], bold=True, size=15, name="Calibri")
    c.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 42

    ws.merge_cells('A2:L2')
    c2 = ws['A2']
    c2.value = f"Generated: {datetime.now().strftime('%d %B %Y, %H:%M')}  |  Threshold TOEFL: ≥{TOEFL_PASS}"
    c2.fill = fill("1A252F")
    c2.font = Font(color="95A5A6", bold=False, size=10, name="Calibri", italic=True)
    c2.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[2].height = 20

    sum_headers = [
        "No", "Kelas", "Semester", "Total\nMahasiswa",
        "Layak\nSertifikat", "File di\nDrive",
        "Belum\ndi Drive", "% Selesai",
        "Nama\nCocok", "Nama\nBeda",
        "Google Drive", "Sheet"
    ]
    for ci, h in enumerate(sum_headers, 1):
        cell = ws.cell(row=3, column=ci)
        cell.value = h
        hdr(cell)
    ws.row_dimensions[3].height = 40

    total_mhs = total_layak = total_yes = total_no = 0
    total_name_match = total_name_mismatch = 0

    for ri, info in enumerate(all_class_data, 1):
        er  = 3 + ri
        bg  = C["ROW_ODD"] if ri % 2 == 0 else C["ROW_EVEN"]
        pct = (info['cert_yes'] / info['layak'] * 100) if info['layak'] > 0 else 0

        total_mhs          += info['total']
        total_layak        += info['layak']
        total_yes          += info['cert_yes']
        total_no           += info['cert_no']
        total_name_match   += info['name_match']
        total_name_mismatch += info['name_mismatch']

        row_vals = [
            ri,
            info['class_name'],
            info['semester'],
            info['total'],
            info['layak'],
            info['cert_yes'],
            info['cert_no'],
            f"{pct:.0f}%",
            info['name_match'],
            info['name_mismatch'],
            info['drive_url'],
            info['sheet_name'],
        ]
        for ci, val in enumerate(row_vals, 1):
            cell = ws.cell(row=er, column=ci)
            if ci == 2:
                cell.value = val
                dat(cell, bg=bg, bold=True, align='left')
            elif ci == 11:   # Google Drive link
                cell.value = "🔗 Buka Drive"
                cell.hyperlink = val
                cell.fill = fill(C["DRIVE_BG"])
                cell.font = Font(color=C["DRIVE_FG"], bold=False, size=10,
                                  underline='single', name="Calibri")
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.border = thin_border
            elif ci == 8:  # Persentase
                cell.value = val
                p = info['cert_yes'] / info['layak'] * 100 if info['layak'] > 0 else 0
                col = C["SCORE_HIGH"] if p >= 80 else (C["SCORE_MED"] if p >= 50 else C["SCORE_LOW"])
                dat(cell, bg=col, bold=True)
            elif ci == 6:   # cert_yes
                cell.value = val
                dat(cell, bg=C["PASS_BG"], bold=True, fg="1E8449")
            elif ci == 7:   # cert_no
                cell.value = val
                dat(cell, bg=C["FAIL_BG"] if val > 0 else bg, bold=(val > 0), fg=(C["CERT_NO"] if val > 0 else "000000"))
            elif ci == 9:   # Nama Cocok
                cell.value = val
                dat(cell, bg=C["PASS_BG"] if val > 0 else bg, bold=True, fg="1E8449")
            elif ci == 10:  # Nama Beda
                cell.value = val
                dat(cell, bg=C["FAIL_BG"] if val > 0 else bg, bold=(val > 0), fg=(C["CERT_NO"] if val > 0 else "000000"))
            else:
                cell.value = val
                dat(cell, bg=bg)
        ws.row_dimensions[er].height = 22

    # Total baris
    sum_row = 3 + len(all_class_data) + 1
    ws.merge_cells(f'A{sum_row}:C{sum_row}')
    tc = ws[f'A{sum_row}']
    tc.value = "TOTAL KESELURUHAN"
    tc.fill = fill(C["SUMMARY_BG"])
    tc.font = Font(color=C["SUMMARY_FG"], bold=True, size=11, name="Calibri")
    tc.alignment = Alignment(horizontal='center', vertical='center')
    tc.border = thin_border
    ws.row_dimensions[sum_row].height = 28

    pct_total = total_yes / total_layak * 100 if total_layak > 0 else 0
    for ci, val in enumerate(
        [total_mhs, total_layak, total_yes, total_no, f"{pct_total:.0f}%",
         total_name_match, total_name_mismatch],
        start=4
    ):
        cell = ws.cell(row=sum_row, column=ci)
        cell.value = val
        cell.fill = fill(C["SUMMARY_BG"])
        cell.font = Font(color=C["SUMMARY_FG"], bold=True, size=11, name="Calibri")
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = thin_border

    col_w = [5, 16, 14, 12, 12, 12, 12, 12, 12, 12, 16, 22]
    for i, w in enumerate(col_w, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = ws['A4']


# ─── MAIN ─────────────────────────────────────────────────────────────────────
def main():
    print("=" * 60)
    print("  MONITORING MAHASISWA TOEFL – Excel Generator")
    print("=" * 60)

    with open(DATA_JSON, "r") as f:
        config = json.load(f)

    print("\n[1] Autentikasi Google Drive...")
    try:
        drive_service = get_drive_service()
        print("    OK – Drive terhubung.")
    except Exception as e:
        print(f"    ERROR: {e}")
        print("    Lanjut tanpa Drive (kolom sertifikat akan N/A).")
        drive_service = None

    print("\n[1b] Load config SIAKAD...")
    full_check = '--full' in sys.argv or '-f' in sys.argv
    siakad_config = load_siakad_config()
    siakad_cache = {}

    if siakad_config:
        print("    OK – Cookie SIAKAD ditemukan.")
        if full_check:
            print("    MODE: FULL CHECK — cek semua mahasiswa dari awal")
            siakad_cache = {}
        else:
            siakad_cache = load_siakad_cache()
            print(f"    Cache: {len(siakad_cache)} NIM ter-cache.")
            print("    Nama 'Cocok' → skip. Nama 'Beda'/'N/A' → re-cek.")
        print("    Cookie SIAKAD bisa expired. Jika semua N/A, update")
        print("    auth/siakad_session.json dari browser.")
    else:
        print("    SKIP – Tidak ada config SIAKAD (kolom nama akan N/A).")
        print("           Buat auth/siakad_session.json untuk mengaktifkan.")

    print(f"\n[2] Memproses {len(config['files'])} kelas...")
    wb = openpyxl.Workbook()
    all_class_data = []

    for file_info in config['files']:
        try:
            info = build_class_sheet(wb, file_info, drive_service,
                                     siakad_config, siakad_cache, full_check)
            all_class_data.append(info)
            print(f"    ✅ {info['class_name']}: {info['total']} mhs, "
                  f"{info['layak']} layak, {info['cert_yes']} sertifikat di Drive, "
                  f"{info['name_match']} nama cocok, {info['name_mismatch']} nama beda")
        except Exception as e:
            print(f"    ❌ ERROR {file_info['filePath']}: {e}")
            import traceback; traceback.print_exc()

    print("\n[3] Membuat sheet Ringkasan...")
    build_summary_sheet(wb, all_class_data)
    wb.move_sheet("RINGKASAN", offset=-len(wb.sheetnames) + 1)

    timestamp   = datetime.now().strftime("%Y%m%d_%H%M")
    output_name = f"Monitoring_TOEFL_{timestamp}.xlsx"
    output_file = os.path.join(OUTPUT_DIR, output_name)
    wb.save(output_file)

    print(f"\n{'=' * 60}")
    print(f"✨ SELESAI!")
    print(f"   File: {output_file}")
    print(f"   Kelas: {len(all_class_data)}")
    print(f"   Total mahasiswa: {sum(d['total'] for d in all_class_data)}")
    print(f"   Total layak sertifikat: {sum(d['layak'] for d in all_class_data)}")
    print(f"   Sertifikat ada di Drive: {sum(d['cert_yes'] for d in all_class_data)}")
    print(f"   Nama cocok SIAKAD: {sum(d['name_match'] for d in all_class_data)}")
    print(f"   Nama beda SIAKAD:  {sum(d['name_mismatch'] for d in all_class_data)}")
    print(f"{'=' * 60}")

    # Simpan cache SIAKAD untuk run berikutnya
    if siakad_config and siakad_cache:
        save_siakad_cache(siakad_cache)
        cached_ok = sum(1 for v in siakad_cache.values()
                        if v.get("last_match") is True)
        print(f"   Cache SIAKAD disimpan: {len(siakad_cache)} NIM "
              f"({cached_ok} cocok, {len(siakad_cache) - cached_ok} bermasalah)")

    # Buka output folder
    os.startfile(OUTPUT_DIR)


if __name__ == "__main__":
    main()

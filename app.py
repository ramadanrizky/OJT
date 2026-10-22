from flask import Flask, request, send_file, send_from_directory, jsonify, render_template, session, redirect, url_for
from flask_cors import CORS
from werkzeug.utils import secure_filename
from werkzeug.security import generate_password_hash, check_password_hash
from pdf2docx import Converter
from docx2pdf import convert

from functools import wraps
# Import tambahan untuk fitur ringkas dokumen
from docx import Document
from sumy.parsers.plaintext import PlaintextParser
from sumy.nlp.tokenizers import Tokenizer
from sumy.summarizers.lex_rank import LexRankSummarizer
# Import tambahan untuk meningkatkan akurasi Bahasa Indonesia
from Sastrawi.Stemmer.StemmerFactory import StemmerFactory
from Sastrawi.StopWordRemover.StopWordRemoverFactory import StopWordRemoverFactory
import datetime
import os
import re
import base64
import requests

# ===================== TAMBAHAN: AUTO-DOWNLOAD DATA NLTK =====================
# Sumy (peringkasan dokumen) butuh data tokenizer NLTK untuk memecah teks jadi
# kalimat. Blok ini otomatis mengunduh data tersebut kalau belum ada, supaya
# tidak perlu jalankan perintah manual tiap pindah komputer atau deploy ke
# server baru.
import nltk

def _pastikan_data_nltk_tersedia():
    paket_dibutuhkan = ['punkt_tab', 'punkt']
    for paket in paket_dibutuhkan:
        try:
            nltk.data.find(f'tokenizers/{paket}')
        except LookupError:
            print(f"[NLTK] Data '{paket}' belum ada, mengunduh otomatis...")
            try:
                nltk.download(paket, quiet=True)
                print(f"[NLTK] Berhasil mengunduh '{paket}'.")
            except Exception as e:
                print(f"[PERINGATAN] Gagal mengunduh data NLTK '{paket}': {e}")
                print("[PERINGATAN] Fitur Ringkas Dokumen mungkin tidak berfungsi.")

_pastikan_data_nltk_tersedia()
# ===============================================================================


# Beri tahu Flask di mana folder template berada (yaitu, direktori saat ini, '.')
app = Flask(__name__, template_folder='.')
CORS(app) # Mengizinkan frontend mengambil data dari backend

# Kunci rahasia untuk mengamankan sesi. Ganti dengan string acak yang kuat.
app.secret_key = 'ganti-dengan-kunci-rahasia-yang-sangat-aman'

# ===================== TAMBAHAN: BATAS UKURAN UPLOAD =====================
# Dokumen/berkas yang diunggah bisa besar. Batas ini mencegah server kehabisan
# memori/disk; sesuaikan bila perlu (contoh: 500 MB).
app.config['MAX_CONTENT_LENGTH'] = 500 * 1024 * 1024
# ===========================================================================

# ===================== DATABASE MYSQL (PENGGANTI FIRESTORE) =====================
# Semua data (User, Tugas, Pengumuman, Jadwal Rapat, Disposisi, dll.)
# sekarang tersimpan di MySQL/MariaDB. Adapter + skema tabel ada di db_mysql.py.
#
# Format data dan API TIDAK berubah: 1 dokumen Firestore = 1 baris MySQL
# (kolom `id` + `data` JSON + kolom kunci untuk pencarian).
#
# WAJIB: pip install PyMySQL
# Konfigurasi lewat environment (default cocok untuk XAMPP):
#   MYSQL_HOST=127.0.0.1  MYSQL_PORT=3306  MYSQL_USER=root
#   MYSQL_PASSWORD=       MYSQL_DB=portal_karyawan
#
# Kembali sementara ke Firestore (rollback): set DATA_BACKEND=firestore
DATA_BACKEND = os.environ.get('DATA_BACKEND', 'mysql').strip().lower()

if DATA_BACKEND == 'firestore':
    import firebase_admin
    from firebase_admin import credentials, firestore as _firestore

    FIREBASE_CRED_FILENAME = 'portal-staf-32f99-firebase-adminsdk-fbsvc-128d85e3f3.json'
    FIREBASE_CRED_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), FIREBASE_CRED_FILENAME)
    try:
        cred = credentials.Certificate(FIREBASE_CRED_PATH)
        firebase_admin.initialize_app(cred)
        fdb = _firestore.client()
        print(f"[OK] Firestore terhubung (mode rollback), memakai kunci: {FIREBASE_CRED_FILENAME}")
    except Exception as e:
        fdb = None
        print(f"[PERINGATAN] Gagal menghubungkan Firestore: {e}")
        print(f"[PERINGATAN] Pastikan file '{FIREBASE_CRED_FILENAME}' ada di folder project ini.")
    SERVER_TIMESTAMP = _firestore.SERVER_TIMESTAMP
    DELETE_FIELD = _firestore.DELETE_FIELD
    ArrayUnion = _firestore.ArrayUnion
    fdb_transactional = _firestore.transactional
else:
    import db_mysql
    fdb = db_mysql.get_db()
    SERVER_TIMESTAMP = db_mysql.SERVER_TIMESTAMP
    DELETE_FIELD = db_mysql.DELETE_FIELD
    ArrayUnion = db_mysql.ArrayUnion
    fdb_transactional = db_mysql.transactional
# =======================================================================================


# ===================== TELEGRAM BOT CONFIGURATION =====================
# Konfigurasi Telegram Bot untuk notifikasi
# Dapatkan BOT_TOKEN dari @BotFather di Telegram
# Chat ID user didapatkan setelah user memulai chat dengan bot
TELEGRAM_BOT_TOKEN = os.environ.get('TELEGRAM_BOT_TOKEN', '8934991541:AAHGxS-qS8IRSgG_LaxwfBJKdQrWBQse3Ys')
TELEGRAM_API_URL = f'https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}'

def send_telegram_message(chat_id, text, parse_mode='HTML'):
    """Kirim pesan ke Telegram chat ID"""
    if not TELEGRAM_BOT_TOKEN or TELEGRAM_BOT_TOKEN == 'YOUR_BOT_TOKEN_HERE':
        print('[TELEGRAM] Bot token belum dikonfigurasi')
        return False
    
    try:
        url = f'{TELEGRAM_API_URL}/sendMessage'
        payload = {
            'chat_id': chat_id,
            'text': text,
            'parse_mode': parse_mode
        }
        response = requests.post(url, json=payload, timeout=10)
        if response.status_code == 200:
            print(f'[TELEGRAM] Pesan terkirim ke chat_id: {chat_id}')
            return True
        else:
            print(f'[TELEGRAM] Gagal kirim: {response.status_code} - {response.text}')
            return False
    except Exception as e:
        print(f'[TELEGRAM] Error: {e}')
        return False

def send_telegram_document(chat_id, file_path, caption=None, timeout=30):
    """Kirim file (docx/pdf/gambar) ke Telegram chat ID."""
    if not TELEGRAM_BOT_TOKEN or TELEGRAM_BOT_TOKEN == 'YOUR_BOT_TOKEN_HERE':
        print('[TELEGRAM] Bot token belum dikonfigurasi')
        return False
    if not file_path or not os.path.exists(file_path):
        print(f'[TELEGRAM] File tidak ditemukan: {file_path}')
        return False
    try:
        url = f'{TELEGRAM_API_URL}/sendDocument'
        with open(file_path, 'rb') as fh:
            files = {'document': (os.path.basename(file_path), fh)}
            data = {'chat_id': chat_id}
            if caption:
                data['caption'] = caption
                data['parse_mode'] = 'HTML'
            response = requests.post(url, data=data, files=files, timeout=timeout)
        if response.status_code == 200:
            print(f'[TELEGRAM] Dokumen terkirim ke chat_id: {chat_id}')
            return True
        print(f'[TELEGRAM] Gagal kirim dokumen: {response.status_code} - {response.text}')
        return False
    except Exception as e:
        print(f'[TELEGRAM] Error kirim dokumen: {e}')
        return False

# ===================== PENANDA WARNA JENIS SURAT DISPOSISI =====================
# Satu sumber kebenaran untuk urgensi surat. Dipakai di caption Telegram
# (emoji, karena Telegram tak mendukung teks berwarna) dan dipadankan di
# frontend web (dot + badge + border baris).
JENIS_MARKER = {
    'sangat segera': {'dot': 'red',    'emoji': '🔴', 'label': 'SANGAT SEGERA', 'pulse': True},
    'segera':        {'dot': 'yellow', 'emoji': '🟡', 'label': 'SEGERA',         'pulse': True},
    'rahasia':       {'dot': 'black',  'emoji': '⬛', 'label': 'RAHASIA',        'pulse': False},
    'dinas':         {'dot': 'blue',   'emoji': '🔵', 'label': 'DINAS',          'pulse': False},
    'undangan':      {'dot': 'green',  'emoji': '🟢', 'label': 'UNDANGAN',       'pulse': False},
    'laporan':       {'dot': 'purple', 'emoji': '🟣', 'label': 'LAPORAN',        'pulse': False},
    'biasa':         {'dot': 'gray',   'emoji': '⚪', 'label': 'BIASA',          'pulse': False},
}

def marker_jenis(jenis):
    """Kembalikan info penanda untuk suatu jenis surat (case-insensitive)."""
    return JENIS_MARKER.get((jenis or '').strip().lower(), JENIS_MARKER['biasa'])
# ===================================================================================


def notifikasi_disposisi_telegram(penerima, info):
    """Kirim notifikasi + file surat disposisi ke Telegram penerima (atasan/user).

    `penerima`: dict user (wajib ada telegram_chat_id bila ingin terkirim).
    `info`: dict berisi perihal, nomor_surat, no_agenda, pengirim_nama,
            file_docx (path), lampiran_pdf (path atau ''), app_link.
    Mengembalikan True bila minimal pesan teks terkirim.
    """
    chat_id = (penerima or {}).get('telegram_chat_id')
    if not chat_id:
        print(f"[TELEGRAM] {(penerima or {}).get('fullname', '?')} belum menghubungkan Telegram, notifikasi disposisi dilewati")
        return False
    mk = marker_jenis(info.get('jenis_surat'))
    judul = f'{mk["emoji"]} <b>DISPOSISI BARU DITERIMA</b>'
    if mk['pulse']:
        judul = f'{mk["emoji"]} <b>DISPOSISI {mk["label"]} — SEGERA DITINDAKLANJUTI</b>'

    def _v(teks):
        # Surat RAHASIA: isi sensitif disamarkan (spoiler) — penerima ketuk untuk melihat
        t = teks or '-'
        if mk['dot'] == 'black':
            t = t.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
            return f'<span class="tg-spoiler">{t}</span>'
        return t

    caption = (
        f'{judul}\n\n'
        f'{mk["emoji"]} <b>Jenis:</b> {mk["label"]}\n'
        f'📝 <b>Perihal:</b> {_v(info.get("perihal"))}\n'
        f'🔢 <b>No. Surat:</b> {_v(info.get("nomor_surat"))}\n'
        f'🗂 <b>No. Agenda:</b> {_v(info.get("no_agenda"))}\n'
        f'👤 <b>Dari:</b> {_v(info.get("pengirim_nama"))}\n'
        f'⏰ <b>Waktu:</b> {datetime.datetime.now().strftime("%d %B %Y %H:%M")}'
    )
    if mk['dot'] == 'black':
        caption += '\n\n🤫 <i>Isi di atas disamarkan — ketuk teks tersamar untuk melihat.</i>'
    if info.get('app_link'):
        caption += f'\n\n<a href="{info["app_link"]}">Buka aplikasi untuk melihat & meneruskan</a>'
    terkirim = False
    # 1) File lembar disposisi (docx) + caption ringkas
    if info.get('file_docx'):
        terkirim = send_telegram_document(chat_id, info['file_docx'], caption=caption) or terkirim
    else:
        terkirim = send_telegram_message(chat_id, caption) or terkirim
    # 2) Lampiran PDF surat undangan (bila ada)
    if info.get('lampiran_pdf'):
        send_telegram_document(chat_id, info['lampiran_pdf'],
                               caption='📎 <b>Lampiran PDF surat undangan</b>')
    return terkirim

def get_cloudflared_url():
    """Ambil URL publik dari cloudflared environment variable atau generate dari request"""
    # Coba dapatkan dari environment variable (cloudflared biasanya set ini)
    url = os.environ.get('CLOUDFLARED_URL') or os.environ.get('NGROK_URL')
    if url:
        # Pastikan URL memiliki protokol
        if not url.startswith(('http://', 'https://')):
            url = f'https://{url}'
        return url
    # Fallback: gunakan host dari request (akan berubah-ubah)
    return None

def send_task_notification_telegram(user_data, task_data, assigned_by_name):
    """Kirim notifikasi tugas baru via Telegram"""
    if not user_data.get('telegram_chat_id'):
        return False
    
    priority_emoji = {'tinggi': '🔴', 'sedang': '🟡', 'rendah': '🟢'}.get(task_data.get('prioritas', 'sedang'), '🟡')
    
    deadline_text = ''
    if task_data.get('deadline'):
        try:
            deadline_dt = datetime.datetime.strptime(task_data['deadline'], '%Y-%m-%d')
            deadline_text = f'\n📅 <b>Deadline:</b> {deadline_dt.strftime("%d %B %Y")}'
        except:
            deadline_text = f'\n📅 <b>Deadline:</b> {task_data["deadline"]}'
    template_text = f'\n📎 <b>Template:</b> {task_data["template"]}' if task_data.get('template') else ''
    
    # Dapatkan URL cloudflared yang sedang aktif
    public_url = get_cloudflared_url()
    
    # Bangun link tugas - jika ada URL publik, gunakan link lengkap
    # jika tidak, gunakan tautan ke aplikasi
    if public_url:
        task_link = f'{public_url}/tugas-saya.html'
    else:
        task_link = 'Aplikasi Portal (buka di browser)'
    
    message = (
        f'📋 <b>TUGAS BARU DITUGASKAN</b>\n\n'
        f'📝 <b>Judul:</b> {task_data.get("judul", "-")}\n'
        f'📄 <b>Deskripsi:</b> {task_data.get("deskripsi", "-") or "-"}\n'
        f'{priority_emoji} <b>Prioritas:</b> {task_data.get("prioritas", "sedang").capitalize()}{deadline_text}{template_text}\n'
        f'👤 <b>Diberikan oleh:</b> {assigned_by_name}\n'
        f'⏰ <b>Waktu:</b> {datetime.datetime.now().strftime("%d %B %Y %H:%M")}\n\n'
        f'<a href="{task_link}">Buka aplikasi untuk melihat detail & mengerjakan tugas</a>'
    )
    
    return send_telegram_message(user_data['telegram_chat_id'], message)


UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'outputs'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

# ===================== FILE HASIL TUGAS (staf -> persetujuan atasan) =====================
TUGAS_UPLOAD_FOLDER = os.path.join(UPLOAD_FOLDER, 'tugas')
os.makedirs(TUGAS_UPLOAD_FOLDER, exist_ok=True)
ALLOWED_TUGAS_EXTENSIONS = {'pdf', 'doc', 'docx', 'xls', 'xlsx', 'ppt', 'pptx', 'txt', 'csv', 'zip', 'rar', 'jpg', 'jpeg', 'png'}
MAX_TUGAS_FILE_MB = 10

def _tugas_file_ok(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_TUGAS_EXTENSIONS
# =======================================================================================

# ==========================================
# MODEL DATABASE (Cetak Biru Tabel)
# ==========================================
# CATATAN: Data (User, Tugas, JadwalRapat, Announcement, Disposisi)
# disimpan di MySQL/MariaDB melalui adapter `db_mysql.py`. Tidak ada lagi
# model SQLAlchemy di aplikasi ini; skema tabel diatur otomatis oleh adapter.


# ==========================================
# DECORATORS UNTUK KEAMANAN
# ==========================================
def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'user_id' not in session:
            # Jika ini adalah permintaan API, kirim error JSON. Jika tidak, alihkan.
            if request.path.startswith('/api/'):
                return jsonify(error="Sesi telah berakhir, silakan login kembali."), 401
            return redirect(url_for('serve_login_page'))
        return f(*args, **kwargs)
    return decorated_function

def admin_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'user_id' not in session:
            if request.path.startswith('/api/'):
                return jsonify(error="Sesi telah berakhir, silakan login kembali."), 401
            return redirect(url_for('serve_login_page'))
        if session.get('user_role') != 'admin':
            if request.path.startswith('/api/'):
                return jsonify(error="Akses admin diperlukan."), 403
            return "Akses Ditolak", 403
        return f(*args, **kwargs)
    return decorated_function

# ===================== AKSES MODUL DISPOSISI: PERSURATAN =====================
# Hanya Admin, Atasan, dan user dengan bagian='persuratan' yang boleh
# menyentuh surat disposisi & jejaknya. User biasa DITOLAK di semua endpoint.
# Admin dapat mengatur hak per-pengguna lewat Kelola Pengguna (flag eksplisit
# disp_pantau / disp_kirim) yang mengalahkan aturan bawaan di bawah ini.
def _hak_disp_efektif(u, role, user_id):
    """Kembalikan (boleh_pantau, boleh_kirim) disposisi untuk seorang user.
    Rantai disposisi berakhir di atasan tingkat 4 (Dansat): tingkat 1-2 &
    4+ hanya memantau; yang boleh mengirim/meneruskan hanya atasan tingkat 3
    (Asintelops/Asminlog, atau atasan tanpa posisi struktur), persuratan,
    dan admin. Flag eksplisit True/False mengalahkan bawaan; admin selalu
    penuh; kirim => pantau."""
    bagian = ((u or {}).get('bagian') or '')
    if role == 'admin':
        return True, True
    pantau = (role == 'atasan') or (role == 'user' and bagian == 'persuratan')
    if role == 'atasan':
        _t = tingkat_of(user_id)
        kirim = (_t == 3 or _t is None)
    elif role == 'user' and bagian == 'persuratan':
        kirim = True
    else:
        kirim = False
    if (u or {}).get('disp_pantau') is True:
        pantau = True
    elif (u or {}).get('disp_pantau') is False:
        pantau = False
    if (u or {}).get('disp_kirim') is True:
        kirim = True
    elif (u or {}).get('disp_kirim') is False:
        kirim = False
    if kirim:
        pantau = True
    return pantau, kirim

def _akses_disposisi():
    try:
        if fdb is None:
            return False
        d = fdb.collection('users').document(session.get('user_id')).get()
        if not d.exists:
            return False
        pantau, _ = _hak_disp_efektif(d.to_dict() or {}, session.get('user_role'), session.get('user_id'))
        return pantau
    except Exception:
        return False

def disposisi_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'user_id' not in session:
            if request.path.startswith('/api/'):
                return jsonify(error="Sesi telah berakhir, silakan login kembali."), 401
            return redirect(url_for('serve_login_page'))
        if not _akses_disposisi():
            if request.path.startswith('/api/'):
                return jsonify(error="Akses modul disposisi khusus Admin, Atasan, atau Staf Persuratan."), 403
            return "Akses modul disposisi khusus Admin, Atasan, atau Staf Persuratan.", 403
        return f(*args, **kwargs)
    return decorated_function
# =========================================================================

# ============ TINGKAT ATASAN DARI DIREKTORI STAF (STRUKTUR) ============
# Tingkat 1 = Dansatsiber (akar, tanpa parent), 2 = Wadansatsiber, dst.
# Dipakai aturan rantai: tingkat 1-2 & 4+ hanya memantau (rantai berakhir di Dansat).
def tingkat_of(user_id):
    """Kembalikan tingkat (1,2,3,...) dari node struktur milik user.
    None bila user tidak tertaut ke struktur (staf biasa/persuratan)."""
    if not user_id:
        return None
    try:
        nodes = {}
        for d in fdb.collection('struktur').stream():
            x = d.to_dict() or {}
            nodes[d.id] = (x.get('parent_id'), x.get('user_id'))
        target = None
        for nid, (_par, _uid) in nodes.items():
            if _uid == user_id:
                target = nid
                break
        if target is None:
            return None
        dalam = 1
        cur = target
        lihat = set()
        while cur and cur not in lihat and dalam <= 20:
            lihat.add(cur)
            par, _ = nodes.get(cur, (None, None))
            if not par or par not in nodes:
                break
            cur = par
            dalam += 1
        return dalam
    except Exception:
        return None


def is_persuratan(user_id):
    """True bila akun adalah staf persuratan (role=user + bagian=persuratan)."""
    if not user_id:
        return False
    try:
        d = fdb.collection('users').document(user_id).get()
        if not d.exists:
            return False
        x = d.to_dict() or {}
        return x.get('role') == 'user' and x.get('bagian') == 'persuratan'
    except Exception:
        return False


def _cek_pin_disposisi():
    """Verifikasi PIN disposisi lapis-dua untuk staf persuratan pada aksi
    sensitif (Ubah/Hapus surat). PIN dikirim via header X-Disposisi-PIN
    (atau field JSON 'pin'). Kembalikan (True, None) bila lolos — atau
    (False, response_403) bila wajib PIN tapi tidak valid. Non-persuratan
    tidak dikenai PIN."""
    saya = session.get('user_id')
    if not is_persuratan(saya):
        return True, None
    pin = (request.headers.get('X-Disposisi-PIN') or '').strip()
    if not pin and request.is_json:
        try:
            pin = str((request.get_json(silent=True) or {}).get('pin') or '').strip()
        except Exception:
            pin = ''
    try:
        d = fdb.collection('users').document(saya).get()
        tersimpan = ((d.to_dict() or {}) if d.exists else {}).get('pin_disposisi_hash') or ''
    except Exception:
        tersimpan = ''
    if not tersimpan:
        return False, (jsonify({"error": "Belum ada PIN disposisi. Atur dulu di halaman Profil."}), 403)
    if not pin or not check_password_hash(tersimpan, pin):
        return False, (jsonify({"error": "PIN disposisi salah."}), 403)
    return True, None


def _boleh_mengirim():
    """Boleh kirim/teruskan disposisi menurut hak efektif (bawaan: admin,
    atasan tingkat 3, persuratan; atasan tingkat 1-2 & 4+ tidak — kecuali
    diberi flag admin)."""
    if session.get('user_role') == 'admin':
        return True
    try:
        if fdb is None:
            return False
        d = fdb.collection('users').document(session.get('user_id')).get()
        if not d.exists:
            return False
        _, kirim = _hak_disp_efektif(d.to_dict() or {}, session.get('user_role'), session.get('user_id'))
        return kirim
    except Exception:
        return False


def _hanya_pantau():
    """True bila sesi dibatasi hanya melihat daftar, Jejak, Rantai, dan
    Konfirmasi Diterima (tanpa unduh/cetak/kirim/teruskan/hapus)."""
    if session.get('user_role') != 'atasan':
        return False
    try:
        if fdb is not None:
            d = fdb.collection('users').document(session.get('user_id')).get()
            if d.exists:
                pantau, kirim = _hak_disp_efektif(d.to_dict() or {}, 'atasan', session.get('user_id'))
                return bool(pantau and not kirim)
    except Exception:
        pass
    _t = tingkat_of(session.get('user_id'))
    return _t is not None and _t <= 2

def _unduh_ditolak(mode_lihat=False):
    """True bila unduh/cetak Word/PDF surat disposisi DITOLAK untuk sesi ini.

    Aturan: atasan pantau-only tingkat 1 & 2 DIIZINKAN mengunduh Word/PDF
    surat dari persuratan (lihat inline maupun unduh file). Selain itu
    (pantau tingkat 4+, dsb.) hanya boleh mode lihat (?lihat=1).
    Bukan pantau => ikut aturan `boleh` biasa (tidak ditolak di sini)."""
    if not _hanya_pantau():
        return False
    if mode_lihat:
        return False
    try:
        return tingkat_of(session.get('user_id')) not in (1, 2)
    except Exception:
        return True
# =========================================================================

# ===================== TAMBAHAN: DECORATOR ATASAN =====================
def atasan_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'user_id' not in session:
            if request.path.startswith('/api/'):
                return jsonify(error="Sesi telah berakhir, silakan login kembali."), 401
            return redirect(url_for('serve_login_page'))
        if session.get('user_role') not in ('atasan', 'admin'):
            if request.path.startswith('/api/'):
                return jsonify(error="Akses atasan diperlukan."), 403
            return "Akses Ditolak", 403
        return f(*args, **kwargs)
    return decorated_function
# =========================================================================

# =========================================================================

# API 1: Konversi PDF ke Word (kualitas maksimal)
PDF2WORD_MAX_MB = 60
PDF2WORD_MAX_PAGES = 200
_OCR_ENGINE = None

def _get_ocr_engine():
    """Mesin OCR (RapidOCR/PP-OCR) dimuat malas sekali saja. Model (~10MB)
    diunduh otomatis saat pertama dipakai lalu di-cache di komputer ini."""
    global _OCR_ENGINE
    if _OCR_ENGINE is None:
        from rapidocr_onnxruntime import RapidOCR
        _OCR_ENGINE = RapidOCR()
    return _OCR_ENGINE

def _pdf2word_ocr(pdf_path, docx_path, dpi=150):
    """Mode OCR: PDF pindaian/gambar diubah jadi teks Word yang bisa diedit.

    Format dipertahankan semirip mungkin dengan aslinya:
    - ukuran font proporsional tinggi kotak OCR (pt = tinggi_px * 72 / dpi)
    - tebal otomatis untuk judul / teks besar-kapital pendek
    - perataan (kiri/tengah/kanan/justify) dari posisi kotak di halaman
    - warna teks disampling dari citra di tengah kotak
    Tiap halaman di-render lalu dibaca RapidOCR; baris-baris berdekatan
    digabung jadi paragraf mengalir, teks besar & pendek jadi judul."""
    import fitz
    import numpy as np
    from docx import Document
    from docx.shared import Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH

    engine = _get_ocr_engine()
    src = fitz.open(pdf_path)
    doc = Document()
    normal = doc.styles['Normal']
    normal.font.name = 'Times New Roman'
    normal.font.size = Pt(12)
    normal.paragraph_format.space_after = Pt(6)
    normal.paragraph_format.line_spacing = 1.15

    def _warna_dari_citra(img, box, lebar_px, tinggi_px):
        """Ambil warna teks (piksel gelap), bukan latar. Hitam/abu -> None
        (biarkan hitam bawaan Word); warna asli (kop merah/biru) dipertahankan."""
        try:
            import numpy as _np
            xs = [p[0] for p in box]; ys = [p[1] for p in box]
            x0 = max(0, int(min(xs))); x1 = min(lebar_px, int(max(xs)))
            y0 = max(0, int(min(ys))); y1 = min(tinggi_px, int(max(ys)))
            if x1 - x0 < 3 or y1 - y0 < 3:
                return None
            dx = max(1, (x1 - x0) // 4); dy = max(1, (y1 - y0) // 4)
            pot = img[y0 + dy:y1 - dy, x0 + dx:x1 - dx]
            if pot.size == 0:
                return None
            px = pot.reshape(-1, pot.shape[-1]).astype(_np.int32)
            # Piksel teks = yang gelap (rata-rata kanal < 120).
            gelap = px[px.mean(axis=1) < 120]
            if len(gelap) < 8:
                return None
            med = _np.median(gelap, axis=0)
            r, g, b = int(med[0]), int(med[1]), int(med[2])
            if max(r, g, b) < 110 and (max(r, g, b) - min(r, g, b)) < 60:
                return None
            return RGBColor(max(0, min(255, r)), max(0, min(255, g)), max(0, min(255, b)))
        except Exception:
            return None

    for pno, page in enumerate(src):
        if pno > 0:
            doc.add_page_break()
        pix = page.get_pixmap(dpi=dpi)
        img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)
        if pix.n == 4:  # RGBA -> RGB agar OCR stabil
            img = img[:, :, :3].copy()
        result, _ = engine(img)
        if not result:
            doc.add_paragraph(f"[Halaman {pno + 1}: tidak ada teks terbaca]").italic = True
            continue
        # Saring kotak watermark raksasa (tinggi >2,5x median & lebar >40%
        # halaman): hasil pindaian ikut terpindai, jangan jadi teks isi.
        _hs0 = sorted(max(p[1] for p in box) - min(p[1] for p in box)
                      for box, text, score in result if (text or '').strip())
        _med0 = _hs0[len(_hs0) // 2] if _hs0 else 20
        _hasil_saring = []
        for box, text, score in result:
            _xs = [p[0] for p in box]
            _hh = max(p[1] for p in box) - min(p[1] for p in box)
            if _hh > _med0 * 2.5 and (max(_xs) - min(_xs)) > pix.width * 0.4:
                continue
            _hasil_saring.append((box, text, score))
        result = _hasil_saring
        if not result:
            continue
        # Urut atas->bawah, gabung baris rapat jadi satu paragraf.
        # Simpan juga kotak & warna tiap baris untuk pewarisan format.
        lines = sorted(
            [(min(p[1] for p in box), max(p[1] for p in box), min(p[0] for p in box),
              max(p[0] for p in box), box, text, float(score))
             for box, text, score in result if (text or '').strip()],
            key=lambda l: (l[0], l[2]))
        heights = [b - a for a, b, _, _, _, _, _ in lines]
        med_h = sorted(heights)[len(heights) // 2] if heights else 20
        skala_pt = 72.0 / dpi  # tinggi piksel -> poin
        med_pt = med_h * skala_pt
        # Gabung HANYA lanjutan paragraf asli: jarak vertikal rapat DAN
        # tepi kiri hampir sejajar (mencegah kolom form tercampur jadi satu).
        paras, cur = [], []
        for top, bottom, left, right, box, text, score in lines:
            if cur:
                gap = top - cur[-1]['bottom']
                same_flow = abs(left - cur[0]['left']) < med_h * 2.0
                if gap < med_h * 0.5 and same_flow:
                    cur.append({'left': left, 'right': right, 'top': top, 'bottom': bottom,
                                'box': box, 'text': text})
                    continue
                paras.append(cur)
            cur = [{'left': left, 'right': right, 'top': top, 'bottom': bottom,
                    'box': box, 'text': text}]
        if cur:
            paras.append(cur)
        doc.add_heading(f"Halaman {pno + 1}", level=2) if len(src) > 1 else None
        for para in paras:
            para = sorted(para, key=lambda d: d['left'])
            text = ' '.join(d['text'] for d in para)
            text = _bersih_xml(_perbaiki_spasi_hilang(text))
            # pecah kata menempel khas OCR (kop kapital tanpa spasi)
            text = ' '.join(_pecah_kata_menempel(w) for w in text.split())
            # --- Wariskan format dari kotak OCR ---
            h_px = [(d['bottom'] - d['top']) for d in para]
            h_med = sorted(h_px)[len(h_px) // 2]
            ukuran_pt = min(28.0, max(8.0, round(h_med * skala_pt * 1.05 * 2) / 2))
            # Warna: sampling baris tertinggi (paling representatif).
            warna = None
            for d in sorted(para, key=lambda d: (d['bottom'] - d['top']), reverse=True):
                warna = _warna_dari_citra(img, d['box'], pix.width, pix.height)
                if warna is not None:
                    break
            # Perataan dari posisi horizontal di halaman.
            l_min = min(d['left'] for d in para); r_maks = max(d['right'] for d in para)
            m_kiri = l_min / pix.width; m_kanan = (pix.width - r_maks) / pix.width
            bentang = (r_maks - l_min) / pix.width
            if abs(m_kiri - m_kanan) < 0.10 and m_kiri > 0.08:
                align = WD_ALIGN_PARAGRAPH.CENTER
            elif m_kanan < 0.05 and m_kiri > 0.25:
                align = WD_ALIGN_PARAGRAPH.RIGHT
            elif bentang > 0.75 and len(para) > 1:
                align = WD_ALIGN_PARAGRAPH.JUSTIFY
            else:
                align = WD_ALIGN_PARAGRAPH.LEFT
            # Judul: pendek + huruf jauh lebih besar dari rata-rata, atau
            # kapital semua yang pendek (kop surat khas dinas).
            big = sum(1 for h in h_px if h > med_h * 1.8)
            kapital_pendek = len(text) < 90 and text.replace(' ', '').isupper()
            if (len(text) < 80 and big >= len(para) / 2) or kapital_pendek:
                p = doc.add_heading(level=3)
                run = p.add_run(text)
                run.bold = True
                run.font.size = Pt(min(28.0, max(ukuran_pt + 2, med_pt + 4)))
                run.font.name = 'Times New Roman'
                if warna is not None:
                    run.font.color.rgb = warna
                p.alignment = align
            else:
                p = doc.add_paragraph()
                run = p.add_run(text)
                run.font.size = Pt(ukuran_pt)
                run.font.name = 'Times New Roman'
                # Tebalkan bila jelas lebih besar dari isi rata-rata.
                if h_med > med_h * 1.5:
                    run.bold = True
                if warna is not None:
                    run.font.color.rgb = warna
                # Jarak baris proporsional ukuran agar mirip aslinya.
                p.paragraph_format.space_after = Pt(4 if ukuran_pt <= 13 else 8)
                p.alignment = align

    doc.core_properties.title = os.path.basename(pdf_path)
    doc.save(docx_path)
    src.close()

# ============ KONFIGURASI LATIH: PDF-ke-Word Flowing v2 ============
# Semua ambang batas (threshold) tata letak dibaca dari
# pelatihan_konversi/konversi_config_v1.json agar bisa DILATIH ulang
# (lihat pelatihan_konversi/latih_layout.py). Nilai di bawah = default.
DEFAULT_KONVERSI = {
    "margin_min_pt": 28.0, "margin_max_pt": 108.0,
    "full_width_ratio": 0.7, "col_overlap_ratio": 0.3,
    "center_maxw_ratio": 0.65, "center_tol_ratio": 0.05,
    "right_tol_pt": 14.0, "right_gap_pt": 40.0,
    "single_right_edge": 0.85, "single_right_x": 0.4,
    "justify_minw_ratio": 0.82, "justify_tol_pt": 14.0,
    "justify_single_full": 0.97,
    "xycut_gap_y": 10.0, "xycut_gap_x": 24.0,
    "ocr_dpi": 150,
    "indent_tol_pt": 4.0, "firstline_min_pt": 9.0,
    "size_tol_pt": 2.0,
    "head_d1": 3.0, "head_d2": 1.5, "head_len1": 120, "head_len2": 80,
    "head_lvl_gap": 5.0, "kop_halaman_len": 40,
}

def _muat_config_konversi():
    import json as _json
    cfg = dict(DEFAULT_KONVERSI)
    try:
        _p = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                          'pelatihan_konversi', 'konversi_config_v1.json')
        with open(_p, encoding='utf-8') as _fh:
            cfg.update(_json.load(_fh))
    except (OSError, ValueError):
        pass
    return cfg


def _kata_inggris():
    """Himpunan kata Inggris (NLTK) untuk deteksi sambung-kata tanpa hubung.
    Diunduh otomatis sekali, lalu di-cache."""
    global _WORDS_EN
    try:
        return _WORDS_EN
    except NameError:
        pass
    try:
        from nltk.corpus import words as _w
        try:
            _WORDS_EN = set(x.lower() for x in _w.words())
        except LookupError:
            import nltk as _nk
            _nk.download('words', quiet=True)
            _WORDS_EN = set(x.lower() for x in _w.words())
    except Exception:
        _WORDS_EN = set()
    return _WORDS_EN


def _perbaiki_spasi_hilang(t):
    """Sambung kata yang menempel akibat spasi hilang (umum di dokumen
    justify & hasil OCR): 'Bogor,22February' -> 'Bogor, 22 February'.
    Token email/URL/slash tidak disentuh."""
    def _tok(m):
        w = m.group(0)
        if '@' in w or '://' in w or '/' in w:
            return w
        w = re.sub(r'(?<=[a-z,;:])(?=\d)', ' ', w)
        w = re.sub(r'(?<=\d)(?=[A-Za-z])', ' ', w)
        w = re.sub(r'(?<=[a-z])(?=[A-Z][a-z])', ' ', w)
        w = re.sub(r'(?<=\d\.)(?=[A-Za-z])', ' ', w)
        return w
    return re.sub(r'\S+', _tok, t)


def _lem_tanpa_spasi(akhir_sebelum, awal_sesudah):
    """Kode slash (SKET/3526, dan/atau) tidak boleh disisipi spasi."""
    return (akhir_sebelum.endswith('/') or awal_sesudah.startswith('/'))


def _bersih_xml(t):
    """Buang karakter kontrol yang merusak XML Word (NULL, dsb)."""
    import re as _re
    return _re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]', '', t or '')


_KOSAKATA_ID = None

def _muat_kosakata():
    """Muat kosakata Indonesia+istilah dinas (pelatihan_konversi/kosakata_id.txt)."""
    global _KOSAKATA_ID
    if _KOSAKATA_ID is None:
        _KOSAKATA_ID = set()
        try:
            _p = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                              'pelatihan_konversi', 'kosakata_id.txt')
            with open(_p, encoding='utf-8') as _fh:
                for _l in _fh:
                    _w = _l.strip().lower()
                    if _w:
                        _KOSAKATA_ID.add(_w)
        except OSError:
            pass
    return _KOSAKATA_ID

_PENDEK_OK = {"di", "ke", "dan", "tak", "non", "pra", "dwi", "tri", "eka",
              "catur", "panca", "sub", "anti", "pasca", "swasta"}

def _pecah_kata_menempel(token):
    """Pecah token ALFA panjang tanpa spasi (khas OCR kop scan), mis.
    MARKASBESARTENTARANASIONALINDONESIA -> MARKAS BESAR TENTARA NASIONAL
    INDONESIA. Segmentasi greedy-longest + backtracking lawan kamus;
    gagal -> kembalikan asli (aman)."""
    import re as _re
    if not token or len(token) < 10 or len(token) > 60:
        return token
    if ' ' in token or not _re.fullmatch(r'[A-Za-z]+', token):
        return token
    vs = _muat_kosakata()
    if not vs:
        return token
    t = token.lower()
    if t in vs:
        return token

    def _ok(w):
        return len(w) >= 3 or w in _PENDEK_OK

    n = len(t)
    memo = {}

    def _bagi(i):
        if i == n:
            return []
        if i in memo:
            return memo[i]
        best = None
        for j in range(min(n, i + 20), i, -1):
            w = t[i:j]
            if not _ok(w) or w not in vs:
                continue
            rest = _bagi(j)
            if rest is not None:
                cand = [w] + rest
                if best is None or len(cand) < len(best):
                    best = cand
                    if len(best) == 1:
                        break
        memo[i] = best
        return best

    parts = _bagi(0)
    if not parts or len(parts) < 2:
        return token
    out, pos = [], 0
    for w in parts:
        out.append(token[pos:pos + len(w)])
        pos += len(w)
    return ' '.join(out)


def _norm_teks_wm(t):
    import re as _re
    return _re.sub(r'\s+', ' ', (t or '')).strip().lower()


def _daftar_teks_watermark(pdf_path):
    """Kumpulkan TEKS watermark seluruh dokumen: baris berhuruf raksasa
    (>=28pt) dan pudar (luminansi >0,7) yang menindih >=1 baris isi.
    Dipakai SEMUA mode (flowing skip di loop; layout/OCR dibersihkan
    pasca-konversi) agar hasil PDF-ke-Word bebas watermark."""
    import fitz
    hasil = set()
    try:
        src = fitz.open(pdf_path)
    except Exception:
        return hasil
    try:
        for page in src:
            data = page.get_text('dict')
            blok = [b for b in data.get('blocks', []) if b.get('type', 0) == 0]
            garis = []
            for b in blok:
                for l in b.get('lines', []):
                    bb = l.get('bbox', (0, 0, 0, 0))
                    garis.append((bb, l))
            for bb, l in garis:
                sps = [s for s in l.get('spans', []) if (s.get('text') or '').strip()]
                if not sps:
                    continue
                mx = max(float(s.get('size', 11)) for s in sps)
                if mx < 28:
                    continue

                def _lum(c):
                    try:
                        c = int(c)
                        return (0.299 * ((c >> 16) & 255) + 0.587 * ((c >> 8) & 255) + 0.114 * (c & 255)) / 255
                    except Exception:
                        return 0.0

                if not all(_lum(s.get('color')) > 0.7 for s in sps):
                    continue
                tindih = sum(1 for ob, _ in garis
                             if not (ob[2] <= bb[0] or ob[0] >= bb[2]
                                     or ob[3] <= bb[1] or ob[1] >= bb[3]))
                if tindih >= 2:
                    t = _norm_teks_wm(''.join(s.get('text', '') for s in sps))
                    if len(t) >= 3:
                        hasil.add(t)
    finally:
        try:
            src.close()
        except Exception:
            pass
    return hasil


def _buang_watermark_docx(docx_path, wm_set):
    """Hapus paragraf/sel berisi teks watermark dari docx hasil konversi
    (dipakai mode layout pdf2docx yang menyalin mentah). Kembalikan jumlah
    yang dibuang."""
    if not wm_set:
        return 0
    from docx import Document as _Doc
    try:
        doc = _Doc(docx_path)
    except Exception:
        return 0
    buang = 0
    for p in list(doc.paragraphs):
        if _norm_teks_wm(p.text) in wm_set:
            try:
                p._p.getparent().remove(p._p)
                buang += 1
            except Exception:
                pass
    try:
        for t in doc.tables:
            for r in t.rows:
                for c in r.cells:
                    for p in list(c.paragraphs):
                        if _norm_teks_wm(p.text) in wm_set:
                            try:
                                p._p.getparent().remove(p._p)
                                buang += 1
                            except Exception:
                                pass
    except Exception:
        pass
    if buang:
        try:
            doc.save(docx_path)
        except Exception:
            pass
    return buang
# =====================================================================

def _xy_urut(rects, K):
    """Urutkan kotak blok ala XY-cut: belah rekursif pada celah besar.
    Menggantikan sortir-y naif + tebakan 2 kolom: tajuk/kolom/catatan
    kaki terbaca sesuai tata letak visual (seperti iLovePDF)."""
    if len(rects) <= 1:
        return list(rects)
    s = sorted(rects, key=lambda r: (r[1], r[0]))
    cg, ci = -1.0, -1
    for i in range(len(s) - 1):
        g = s[i + 1][1] - s[i][3]
        if g > cg:
            cg, ci = g, i
    if cg > K["xycut_gap_y"]:
        atas = [r for r in s if r[1] < s[ci + 1][1]]
        id_atas = set(map(id, atas))
        bawah = [r for r in s if id(r) not in id_atas]
        if atas and bawah:
            return _xy_urut(atas, K) + _xy_urut(bawah, K)
    t = sorted(s, key=lambda r: (r[0], r[1]))
    cg, ci = -1.0, -1
    for i in range(len(t) - 1):
        g = t[i + 1][0] - t[i][2]
        if g > cg:
            cg, ci = g, i
    if cg > K["xycut_gap_x"]:
        kiri = [r for r in t if r[0] < t[ci + 1][0]]
        id_kiri = set(map(id, kiri))
        kanan = [r for r in t if id(r) not in id_kiri]
        if kiri and kanan:
            return _xy_urut(kiri, K) + _xy_urut(kanan, K)
    return s


def _pdf2word_flowing_v2(pdf_path, docx_path, cfg=None):
    """Mode 'Teks Mengalir' v2 — kualitas setara konverter iLovePDF.

    Beda dari v1 (_pdf2word_flowing):
    - Paragraf digabung lintas baris visual (teks BISA reflow saat diedit,
      bukan satu paragraf per baris). Tanda hubung akhir baris disambung.
    - Urutan baca sadar 2 kolom (kolom kiri habis dulu, baru kanan).
    - Tabel & gambar disisip di POSISI alur baca (bukan dibuang ke akhir).
    - Font/nama/ukuran/italic asli PDF dipertahankan per-run.
    - Baris nomor halaman & kop berulang tidak ikut tercetak tiap halaman.
    """
    import fitz
    from docx import Document
    from docx.shared import Pt, Inches, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
    from docx.oxml.ns import qn
    import re
    import tempfile as _tf

    K = _muat_config_konversi()
    if cfg:
        K.update(cfg)

    _FONT_GANTI = {'helvetica': 'Arial', 'times': 'Times New Roman',
                   'courier': 'Courier New', 'symbol': 'Symbol'}

    def bersih_font(nama):
        n = (nama or '').split('+')[-1].split(',')[0].strip()
        n = re.sub(r'[-,](Bold|Italic|BoldItalic|Oblique|Obli|Regular|Regu|Roma|Ital|Normal|MT|PSMT)+$', '', n, flags=re.I)
        n = n.replace('-', ' ').strip()
        low = n.lower().replace(' ', '')
        _MAP = {'helvetica': 'Arial', 'arial': 'Arial', 'arialmt': 'Arial',
                'nimbussanl': 'Arial',
                'times': 'Times New Roman', 'timesnewroman': 'Times New Roman',
                'timesnewromanps': 'Times New Roman',
                'timesnewromanpsmt': 'Times New Roman', 'nimbusrom': 'Times New Roman',
                'courier': 'Courier New', 'couriernew': 'Courier New',
                'nimbusmon': 'Courier New',
                'palladiol': 'Palatino Linotype', 'urwpalladiol': 'Palatino Linotype',
                'centuryschl': 'Century Schoolbook',
                'bookmanl': 'Bookman Old Style', 'urwbookmanl': 'Bookman Old Style',
                'charter': 'Charter', 'itccharter': 'Charter',
                'dingbats': 'Wingdings', 'standardsyml': 'Symbol', 'symbol': 'Symbol',
                'cmsy10': 'Cambria Math', 'eurm10': 'Cambria Math', 'eurb10': 'Cambria Math',
                'msbm10': 'Cambria Math', 'cmr10': 'Times New Roman', 'cmbx10': 'Times New Roman'}
        return _MAP.get(low, n or 'Calibri')

    def _fmt_key(s):
        try:
            _sz = round(float(s.get('size', 0)), 1)
        except Exception:
            _sz = 0
        try:
            _cl = int(s.get('color') or 0)
        except Exception:
            _cl = 0
        return (_sz, s.get('flags', 0), bersih_font(s.get('font', '')), _cl)

    def tambah_run(par, teks, span, body_size):
        run = par.add_run(teks)
        flags = span.get('flags', 0)
        run.bold = bool(flags & 16)
        run.italic = bool(flags & 2)
        try:
            if span.get('_sup'):
                run.font.superscript = True
            elif span.get('_sub'):
                run.font.subscript = True
        except Exception:
            pass
        try:
            # ukuran ASLI dipertahankan (batas longgar 5-72pt agar display
            # type besar tidak terpotong seperti sebelumnya)
            run.font.size = Pt(max(5.0, min(round(float(span.get('size', body_size)), 1), 72.0)))
        except Exception:
            pass
        try:
            run.font.name = bersih_font(span.get('font', ''))
        except Exception:
            pass
        try:
            warna = span.get('color')
            if warna is not None and int(warna) != 0:
                c = int(warna)
                run.font.color.rgb = RGBColor((c >> 16) & 255, (c >> 8) & 255, c & 255)
        except Exception:
            pass
        return run

    def _pecah_sel(spans, ambang):
        """Pecah spans sebaris jadi sel-sel berdasar celah horizontal."""
        _ss = sorted(spans, key=lambda s: s.get('bbox', (0, 0, 0, 0))[0])
        sel = [[_ss[0]]]
        for s in _ss[1:]:
            _b0 = s.get('bbox', (0, 0, 0, 0))
            _b1 = sel[-1][-1].get('bbox', (0, 0, 0, 0))
            if _b0[0] - _b1[2] > ambang:
                sel.append([s])
            else:
                sel[-1].append(s)
        return sel

    def _baris_jadi_tabel(baris, body_size):
        """baris = list[(line, spans)]. Kembalikan rows bila pola kolom
        konsisten (>=2 kolom sama di semua baris, sejajar <=12pt)."""
        _ambang = max(15.0, float(body_size) * 1.4)
        baris_sel = [_pecah_sel(sp, _ambang) for _, sp in baris]
        if any(len(s) < 1 for s in baris_sel):
            return None
        nkol = len(baris_sel[0])
        if nkol < 2 or any(len(s) != nkol for s in baris_sel):
            return None
        for c in range(nkol):
            _xs = [s[0].get('bbox', (0, 0, 0, 0))[0] for s in (r[c] for r in baris_sel)]
            if max(_xs) - min(_xs) > 12.0:
                return None
        return [[' '.join(s.get('text', '').strip() for s in c).strip()
                 for c in sel] for sel in baris_sel]

    def _blok_jadi_tabel(doc, blok, body_size):
        """Ubah blok teks kolom-sejajar TANPA garis menjadi tabel Word
        (umum di CV/laporan). Kembalikan True bila ditulis sebagai tabel."""
        garis = []
        for l in blok.get('lines', []):
            _sp = [s for s in l.get('spans', []) if (s.get('text') or '').strip()]
            if _sp:
                garis.append((l, _sp))
        if len(garis) < 3:
            return False
        rows = _baris_jadi_tabel(garis, body_size)
        if not rows or not any(any(v for v in r) for r in rows):
            return False
        tulis_tabel(doc, rows)
        return True

    def _scan_tabel_halaman(blok_teks, body_size, page_cx, page_w):
        """Tabel borderless lintas-blok (satu baris = satu blok, umum hasil
        insert per-baris). Kembalikan (daftar_tabel, id_garis_pakai).
        daftar_tabel = list[(y0, rows)]; hanya baris se-kolom & rapat."""
        semua = []
        for b in blok_teks:
            for l in b.get('lines', []):
                _sp = [s for s in l.get('spans', []) if (s.get('text') or '').strip()]
                if _sp:
                    _bb = l.get('bbox', (0, 0, 0, 0))
                    semua.append((_bb[1], l, _sp))
        semua.sort(key=lambda e: e[0])
        _ambang = max(15.0, float(body_size) * 1.4)
        tabel, pakai = [], set()
        i = 0
        while i < len(semua):
            _y, _l, _sp = semua[i]
            _sel = _pecah_sel(_sp, _ambang)
            if len(_sel) < 2:
                i += 1
                continue
            _nkol = len(_sel)
            _run = [(semua[i][1], semua[i][2])]
            _xs0 = [_sel[c][0].get('bbox', (0, 0, 0, 0))[0] for c in range(_nkol)]
            j = i + 1
            while j < len(semua):
                _s2 = _pecah_sel(semua[j][2], _ambang)
                if len(_s2) != _nkol:
                    break
                _xs = [_s2[c][0].get('bbox', (0, 0, 0, 0))[0] for c in range(_nkol)]
                if any(abs(a - b) > 12.0 for a, b in zip(_xs, _xs0)):
                    break
                _gap = semua[j][0] - semua[j - 1][0]
                if _gap > float(body_size) * 2.5:
                    break
                _run.append((semua[j][1], semua[j][2]))
                j += 1
            if len(_run) >= 3:
                rows = [[' '.join(s.get('text', '').strip() for s in c).strip()
                         for c in _pecah_sel(sp, _ambang)] for _, sp in _run]
                if any(any(v for v in r) for r in rows):
                    _bbs = [l.get('bbox', (0, 0, 0, 0)) for l, _ in _run]
                    _x0 = min(b[0] for b in _bbs)
                    _y0 = min(b[1] for b in _bbs)
                    _x1 = max(b[2] for b in _bbs)
                    _y1 = max(b[3] for b in _bbs)
                    tabel.append((_x0, _y0, _x1, _y1, rows))
                    for l, _ in _run:
                        pakai.add(id(l))
                i = j
            else:
                i += 1
        return tabel, pakai

    def tulis_tabel(doc, rows):
        if not rows or not rows[0]:
            return
        nkol = len(rows[0])
        t = doc.add_table(rows=len(rows), cols=nkol)
        t.style = 'Table Grid'
        for ri, row in enumerate(rows):
            for ci in range(nkol):
                val = row[ci] if ci < len(row) else None
                cell = t.cell(ri, ci)
                cell.text = '' if val is None else _bersih_xml(str(val).strip())
                for par in cell.paragraphs:
                    par.style = doc.styles['Normal']
                if ri == 0:
                    for par in cell.paragraphs:
                        for r in par.runs:
                            r.bold = True
                    tcPr = cell._tc.get_or_add_tcPr()
                    shd = tcPr.makeelement(qn('w:shd'), {
                        qn('w:val'): 'clear', qn('w:color'): 'auto', qn('w:fill'): 'D9E2F3'})
                    tcPr.append(shd)
        _pemisah = doc.add_paragraph()
        _pemisah.paragraph_format.space_after = Pt(0)
        _pemisah.paragraph_format.space_before = Pt(0)

    def tulis_tabel_polos(doc, rows):
        """Tabel 2 kolom TANPA garis & tanpa kepala berwarna untuk pasangan
        entri-tahun CV (kolom kanan tahun). Sel memakai run berformat
        (tebal/warna/font asli). Beda dari tulis_tabel bergaris."""
        if not rows or not rows[0]:
            return
        nkol = max(len(r) for r in rows)
        t = doc.add_table(rows=len(rows), cols=nkol)
        try:
            t.style = doc.styles['Normal Table']
        except Exception:
            pass
        try:
            from docx.oxml import OxmlElement as _Ox
            _tbl = t._tbl
            _tblPr = _tbl.tblPr if _tbl.tblPr is not None else _Ox('w:tblPr')
            _bdr = _Ox('w:tblBorders')
            for _bn in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
                _be = _Ox(f'w:{_bn}')
                for _k, _v in (('w:val', 'nil'), ('w:sz', '0'),
                               ('w:space', '0'), ('w:color', 'auto')):
                    _be.set(qn(_k), _v)
                _bdr.append(_be)
            _tblPr.append(_bdr)
        except Exception:
            pass
        for ri, row in enumerate(rows):
            for ci in range(nkol):
                val = row[ci] if ci < len(row) else ('', (False, None, 'Calibri'))
                _tx, _ft = val if isinstance(val, tuple) else (val, (False, None, 'Calibri'))
                _bd, _wr, _fn = _ft if isinstance(_ft, tuple) else (False, None, 'Calibri')
                cell = t.cell(ri, ci)
                par = cell.paragraphs[0]
                par.style = doc.styles['Normal']
                if ci == 1:
                    par.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                _rn = par.add_run(_bersih_xml(str(_tx).strip()))
                _rn.bold = bool(_bd)
                try:
                    _rn.font.name = _fn or 'Calibri'
                    if _wr is not None:
                        _c = int(_wr)
                        _rn.font.color.rgb = RGBColor((_c >> 16) & 255, (_c >> 8) & 255, _c & 255)
                except Exception:
                    pass

    def _pasang_tahun(blok_teks):
        """Pasangan entri-tahun CV: baris tahun di zona kanan + baris entri
        sejajar y-nya digabung jadi baris tabel [entri, tahun].
        Sel membawa format asli (tebal/warna/font) + peluru vektor bila ada.
        Kembalikan (items_tabel, id_garis_pakai). Run >=2 pasangan rapat."""
        def _fmt_sel(_sps):
            _nb = [s for s in _sps if (s.get('text') or '').strip()]
            if not _nb:
                return (False, None, 'Calibri')
            _bd = all(('bold' in (s.get('font', '').lower())) or (s.get('flags', 0) & 16) for s in _nb)
            _cl = {s.get('color') for s in _nb}
            _wr = None
            if len(_cl) == 1:
                _c = next(iter(_cl))
                if _c not in (None, 0):
                    _wr = int(_c)
            try:
                _fn = bersih_font(_nb[0].get('font', ''))
            except Exception:
                _fn = 'Calibri'
            return (_bd, _wr, _fn)

        _TAHUN = re.compile(r'^\d{4}(?:\s*[–—-]\s*(?:\d{4}|[Pp]resent|[Kk]ini|[Ss]ekarang))?$')
        garis = []
        for b in blok_teks:
            for l in b.get('lines', []):
                _sp = [s for s in l.get('spans', []) if (s.get('text') or '').strip()]
                if not _sp:
                    continue
                _bb = l.get('bbox', (0, 0, 0, 0))
                _t = ' '.join(s.get('text', '') for s in _sp).strip()
                if not _t:
                    continue
                garis.append((_bb[1], _bb[0], _bb[2], _bb[3], _t, id(l),
                              max(float(s.get('size', body_size)) for s in _sp),
                              list(_sp), id(l) in _peluru_v))
        if not garis:
            return [], set()
        _zona = m_kiri_pt + lebar_konten_pt * 0.60
        _tahun = [g for g in garis if _TAHUN.match(g[4]) and g[1] >= _zona]
        if not _tahun:
            return [], set()
        _entri = [g for g in garis
                  if not _TAHUN.match(g[4])
                  and not re.fullmatch(r'[\u2022\-\*\u00b7\ufffd\uf0b7\u25cf\u25aa\s]{1,4}', g[4])]

        def _sel(_g):
            _tx = ('• ' if _g[8] else '') + _g[4]
            return (_tx, _fmt_sel(_g[7]))

        def _baris(_run):
            return [[_sel(_e), _sel(_y)] for _, _e, _y in _run]

        pasang = []
        for _y in _tahun:
            _c = [e for e in _entri
                  if abs(e[0] - _y[0]) <= 6.0 and e[2] <= _y[1] - 12.0 and e[1] < _y[1]]
            if not _c:
                continue
            _e = min(_c, key=lambda e: abs(e[0] - _y[0]))
            pasang.append((_y[0], _e, _y))
        pasang.sort(key=lambda p: p[0])
        items_tb, pakai = [], set()
        _run = []
        for _p in pasang:
            if _run and (_p[0] - _run[-1][0] > float(body_size) * 3.5
                         or abs(_p[1][1] - _run[-1][1][1]) > 12.0):
                if len(_run) >= 2:
                    _bbx = [(_e[1], _e[0], _e[2], _e[3]) for _, _e, _ in _run] + \
                           [(_y[1], _y[0], _y[2], _y[3]) for _, _, _y in _run]
                    items_tb.append((min(b[0] for b in _bbx), min(b[1] for b in _bbx),
                                     max(b[2] for b in _bbx), max(b[3] for b in _bbx),
                                     'tabel_polos', _baris(_run)))
                    for _, _e, _y in _run:
                        pakai.add(_e[5])
                        pakai.add(_y[5])
                _run = []
            _run.append(_p)
        if len(_run) >= 2:
            _bbx = [(_e[1], _e[0], _e[2], _e[3]) for _, _e, _ in _run] + \
                   [(_y[1], _y[0], _y[2], _y[3]) for _, _, _y in _run]
            items_tb.append((min(b[0] for b in _bbx), min(b[1] for b in _bbx),
                             max(b[2] for b in _bbx), max(b[3] for b in _bbx),
                             'tabel_polos', _baris(_run)))
            for _, _e, _y in _run:
                pakai.add(_e[5])
                pakai.add(_y[5])
        return items_tb, pakai

    def tulis_gambar(doc, src, xref, lebar_inci, lebar_asli_pt=None,
                       x_pt=None, y_pt=None, konten_pt=None, tinggi_asli_pt=None):
        """Sematkan gambar SEUKURAN tampilan di PDF (bukan dimaksimalkan),
        agar tata letak & jumlah halaman sama seperti sumber.
        Ukuran & aspek mengikuti kotak tampil PDF (rect): bila rect diketahui,
        lebar/tinggi rect MENANG atas estimasi lain (PDF bisa melebarkan
        gambar tak-proporsional terhadap piksel mentahnya).
        Gambar kecil melayang (floating) di posisi absolut PDF sehingga tidak
        memakan ruang alir vertikal; gambar lebar tetap inline."""
        tmp_path = None
        try:
            pix = fitz.Pixmap(src, xref)
            if pix.n - pix.alpha > 3:
                pix = fitz.Pixmap(fitz.csRGB, pix)
            import os as _os
            fd, tmp_path = _tf.mkstemp(suffix='.png')
            _os.close(fd)
            pix.save(tmp_path)
            if lebar_asli_pt:
                # ukuran tampil PDF mengikat: hormati rect, batasi halaman
                _w = max(0.5, min(lebar_asli_pt / 72, 6.5))
            else:
                _w = lebar_inci
                _w = max(0.5, min(_w, 6.5))
            if tinggi_asli_pt and lebar_asli_pt:
                _h_inci = _w * (tinggi_asli_pt / max(lebar_asli_pt, 1))
            else:
                _h_inci = _w * pix.height / max(pix.width, 1)
            _melayang = (konten_pt and x_pt is not None and y_pt is not None
                         and _w < konten_pt / 72 * 0.6)
            if _melayang:
                _tambah_gambar_melayang(doc, tmp_path, x_pt, y_pt, _w, _h_inci)
            else:
                doc.add_picture(tmp_path, width=Inches(_w), height=Inches(_h_inci))
        except Exception:
            pass
        finally:
            try:
                import os as _os2
                if tmp_path and _os2.path.exists(tmp_path):
                    _os2.unlink(tmp_path)
            except OSError:
                pass

    def _tambah_gambar_melayang(doc, gbr_path, x_pt, y_pt, w_inci, h_inci,
                                 belakang=False):
        """Sisip gambar floating (anchor absolut thd halaman) via OXML.
        Jangkar ditempel ke paragraf terakhir yang ada (bukan paragraf baru)
        agar tidak menambah baris kosong yang menggeser tata letak.
        Ukuran dipaksa mengikuti rect PDF (lebar+tinggi) agar aspek sama
        persis seperti sumber. belakang=True -> di belakang teks (watermark)."""
        from docx.shared import Pt as _Pt
        if len(doc.paragraphs) == 0:
            _p0 = doc.add_paragraph()
            _p0.paragraph_format.space_after = _Pt(0)
            _p0.paragraph_format.space_before = _Pt(0)
            if _p0.runs:
                _p0.runs[0].font.size = _Pt(1)
            else:
                _p0.add_run().font.size = _Pt(1)
        _p = doc.paragraphs[-1]
        _run = _p.add_run()
        if h_inci:
            _run.add_picture(gbr_path, width=Inches(w_inci), height=Inches(h_inci))
        else:
            _run.add_picture(gbr_path, width=Inches(w_inci))
        _ubah_ke_jangkar(_run, x_pt, y_pt, belakang=belakang)

    def _ubah_ke_jangkar(_run, x_pt, y_pt, belakang=False):
        """Ubah gambar inline dalam run menjadi anchor absolut halaman."""
        from docx.oxml import OxmlElement
        try:
            _inline = _run._r.xpath('.//wp:inline')[0]
        except Exception:
            return False
        _anchor = OxmlElement('wp:anchor')
        for _k, _v in (('distT', '0'), ('distB', '0'), ('distL', '0'),
                       ('distR', '0'), ('simplePos', '0'),
                       ('relativeHeight', '1'),
                       ('behindDoc', '1' if belakang else '0'),
                       ('locked', '0'), ('layoutInCell', '1'),
                       ('allowOverlap', '1')):
            _anchor.set(_k, _v)
        _sp = OxmlElement('wp:simplePos')
        _sp.set('x', '0')
        _sp.set('y', '0')
        _ph = OxmlElement('wp:positionH')
        _ph.set('relativeFrom', 'page')
        _po = OxmlElement('wp:posOffset')
        _po.text = str(int(x_pt * 12700))
        _ph.append(_po)
        _pv = OxmlElement('wp:positionV')
        _pv.set('relativeFrom', 'page')
        _po2 = OxmlElement('wp:posOffset')
        _po2.text = str(int(y_pt * 12700))
        _pv.append(_po2)
        if not belakang:
            _wrap = OxmlElement('wp:wrapSquare')
            _wrap.set('wrapText', 'bothSides')
        else:
            _wrap = None
        _anchor.append(_sp)
        _anchor.append(_ph)
        _anchor.append(_pv)
        _kids = list(_inline)
        for _ch in _kids:
            # urutan skema wp:anchor: ... extent, [wrap], docPr, ... :
            # sisipkan wrap tepat sebelum docPr
            if _wrap is not None and _ch.tag.endswith('}docPr'):
                _anchor.append(_wrap)
                _wrap = None
            _anchor.append(_ch)
        if _wrap is not None:
            _anchor.append(_wrap)
        _inline.getparent().replace(_inline, _anchor)
        return True

    # (_gambar_watermark DIHAPUS: watermark dinonaktifkan total.)

    _NO_HAL = re.compile(r'^[-–—\s]*\d{1,4}[-–—\s]*$')

    src = fitz.open(pdf_path)

    def _norm_kop(t):
        return re.sub(r'\s+', ' ', (t or '')).strip().lower()

    src = fitz.open(pdf_path)
    doc = Document()
    # --- Deteksi kop berulang: blok teks teratas yang SAMA di >=2 halaman
    # --- dipindah ke header Word (tidak mengotori tiap halaman hasil)
    kop_set, kop_asli = set(), ""
    try:
        if src.page_count >= 2:
            _hitung = {}
            _contoh = {}
            for _pg in src:
                _bloks = [b for b in _pg.get_text('blocks')
                          if len(b) >= 7 and b[6] == 0 and (b[4] or '').strip()]
                if not _bloks:
                    continue
                _atas = min(_bloks, key=lambda b: b[1])
                _nk = _norm_kop(_atas[4])
                if len(_nk) >= 8:
                    _hitung[_nk] = _hitung.get(_nk, 0) + 1
                    _contoh.setdefault(_nk, _atas[4].strip().split('\n')[0][:120])
            for _nk, _c in _hitung.items():
                if _c >= 2:
                    kop_set.add(_nk)
                    kop_asli = kop_asli or _contoh.get(_nk, '')
            if kop_set and kop_asli:
                try:
                    doc.sections[0].header.paragraphs[0].text = kop_asli
                except Exception:
                    pass
    except Exception:
        kop_set = set()
    normal = doc.styles['Normal']
    normal.font.name = 'Calibri'
    normal.font.size = Pt(11)
    normal.paragraph_format.space_after = Pt(6)
    normal.paragraph_format.line_spacing = 1.15
    for level, size in ((1, 16), (2, 14), (3, 12)):
        hs = doc.styles[f'Heading {level}']
        hs.font.name = 'Calibri'
        hs.font.size = Pt(size)
        hs.font.bold = True

    # --- Kalibrasi margin dari tepi konten aktual (gaya iLovePDF) ---
    # Persentil-5 x0 garis (margin kiri) & persentil-95 x1 (margin kanan):
    # kebal terhadap blok menepi (tanda tangan, tembusan) yang mengacaukan
    # median. Dijepit agar outlier tepi (watermark) tidak merusak.
    try:
        _lx, _rx, _tp = [], [], []
        _pw = 595.0
        for _pg in src:
            _pw = _pg.rect.width
            for _b in _pg.get_text('blocks'):
                if len(_b) < 7 or _b[6] != 0:
                    continue  # tuple blocks: (x0,y0,x1,y1,teks,no,tipe); tipe 0 = teks
                _t = (_b[4] if len(_b) > 4 else '') or ''
                _l1 = _t.strip().split('\n')[0] if _t.strip() else ''
                if _l1 and (_NO_HAL.fullmatch(_l1) or (
                        len(_l1) < K["kop_halaman_len"] and _l1.lower().startswith(('halaman ', 'page ')))):
                    continue
                _lx.append(_b[0]); _rx.append(_b[2]); _tp.append(_b[1])
        def _pct(v, q, dflt):
            if not v:
                return dflt
            s = sorted(v)
            return s[min(len(s) - 1, int(len(s) * q))]
        m_kiri_pt = min(max(_pct(_lx, 0.05, 54.0), K["margin_min_pt"]), K["margin_max_pt"])
        m_kanan_pt = min(max(_pw - _pct(_rx, 0.95, _pw - 54.0), K["margin_min_pt"]), K["margin_max_pt"])
        # margin atas/bawah dari tepi konten EKSTREM (min/maks, bukan median):
        # median/persentil membuang isi tepi (kop/paraf) -> ruang susut ->
        # luapan halaman. Nomor halaman sudah dikecualikan dari daftar.
        m_atas_pt = min(max(min(_tp) if _tp else 54.0, K["margin_min_pt"]), K["margin_max_pt"])
        _ph = src[0].rect.height if len(src) else 842.0
        try:
            _yb = []
            for b in src[0].get_text('blocks'):
                if len(b) < 7 or b[6] != 0:
                    continue
                _t0 = (b[4] or '').strip().split('\n')[0] if (b[4] or '').strip() else ''
                if _t0 and (_NO_HAL.fullmatch(_t0) or (
                        len(_t0) < K["kop_halaman_len"] and _t0.lower().startswith(('halaman ', 'page ')))):
                    continue  # nomor halaman kaki jangan tentukan margin bawah
                _yb.append(b[3])
            _m_bawah = _ph - max(_yb) if _yb else m_atas_pt
        except Exception:
            _m_bawah = m_atas_pt
        m_bawah_pt = min(max(_m_bawah, K["margin_min_pt"]), K["margin_max_pt"])
        _sec = doc.sections[0]
        _sec.left_margin = Inches(m_kiri_pt / 72)
        _sec.right_margin = Inches(m_kanan_pt / 72)
        _sec.top_margin = Inches(m_atas_pt / 72)
        _sec.bottom_margin = Inches(m_bawah_pt / 72)
        lebar_konten_pt = _pw - m_kiri_pt - m_kanan_pt
    except Exception:
        m_kiri_pt, m_kanan_pt, m_atas_pt, lebar_konten_pt = 54.0, 54.0, 54.0, 487.0

    for pno, page in enumerate(src):
        # Ukuran + orientasi halaman mengikuti PDF asli (A4/folio/landscape),
        # bukan default Letter-portrait template Word.
        _pw_pt, _ph_pt = page.rect.width, page.rect.height
        if pno == 0:
            try:
                _sec0 = doc.sections[0]
                _sec0.page_width = Pt(_pw_pt)
                _sec0.page_height = Pt(_ph_pt)
            except Exception:
                pass
            _uk_sebelum = (_pw_pt, _ph_pt)
        elif abs(_pw_pt - _uk_sebelum[0]) > 2 or abs(_ph_pt - _uk_sebelum[1]) > 2:
            try:
                from docx.enum.section import WD_SECTION
                _secn = doc.add_section(WD_SECTION.NEW_PAGE)
                _secn.page_width = Pt(_pw_pt)
                _secn.page_height = Pt(_ph_pt)
                _secn.left_margin = Inches(m_kiri_pt / 72)
                _secn.right_margin = Inches(m_kanan_pt / 72)
                _secn.top_margin = Inches(m_atas_pt / 72)
                _secn.bottom_margin = Inches(m_atas_pt / 72)
            except Exception:
                doc.add_page_break()
            _uk_sebelum = (_pw_pt, _ph_pt)
        else:
            _pb = doc.add_paragraph()
            _rpb = _pb.add_run()
            _rpb.add_break(WD_BREAK.PAGE)
            _pb.paragraph_format.space_after = Pt(0)
            _pb.paragraph_format.space_before = Pt(0)
        data = page.get_text('dict')
        blok_teks = [b for b in data.get('blocks', []) if b.get('type', 0) == 0]
        if not blok_teks and not list(page.find_tables()) and not page.get_images(full=True):
            continue

        # Ukuran font tubuh = modus TERTIMBANG jumlah karakter (stabil walau
        # ada catatan kaki kecil / judul besar; modus span mentah seri).
        from collections import Counter as _C
        sizes, fonts = _C(), _C()
        for b in blok_teks:
            for l in b.get('lines', []):
                for s in l.get('spans', []):
                    _tx = s.get('text', '').strip()
                    if _tx:
                        sizes[round(s.get('size', 11))] += len(_tx)
                        fonts[bersih_font(s.get('font', ''))] += len(_tx)
        body_size = sizes.most_common(1)[0][0] if sizes else 11
        if fonts:
            try:
                normal.font.name = fonts.most_common(1)[0][0]
            except Exception:
                pass
        page_cx = page.rect.width / 2
        # Kata per halaman (sekali saja) untuk uji regang-spasi justify
        try:
            _kata_hal = page.get_text('words')
        except Exception:
            _kata_hal = []

        def _regang_spasi(y_atas, y_bawah, ambang=1.30):
            """True bila spasi antar-kata direnggangkan (ciri justify asli),
            bukan spasi normal teks rata kiri."""
            _w = [w for w in _kata_hal
                  if w[1] >= y_atas - 2 and w[3] <= y_bawah + 2 and str(w[4]).strip()]
            if len(_w) < 4:
                return False
            _w.sort(key=lambda w: w[0])
            _teks = sum(w[2] - w[0] for w in _w)
            _bentang = max(w[2] for w in _w) - min(w[0] for w in _w)
            _celah = _bentang - _teks
            if _celah <= 0:
                return False
            try:
                _normal = fitz.get_text_length(' ', fontsize=float(body_size)) or 1.0
            except Exception:
                _normal = 1.0
            return (_celah / max(len(_w) - 1, 1)) > _normal * ambang

        # --- Deteksi 2 kolom: blok menyebar di kiri & kanan dengan tumpang vertikal
        kiri = [b for b in blok_teks if (b['bbox'][0] + b['bbox'][2]) / 2 < page_cx]
        kanan = [b for b in blok_teks if (b['bbox'][0] + b['bbox'][2]) / 2 >= page_cx]
        dua_kolom = False
        if len(kiri) >= 2 and len(kanan) >= 2:
            atas_kiri = min(b['bbox'][1] for b in kiri)
            bawah_kiri = max(b['bbox'][3] for b in kiri)
            atas_kanan = min(b['bbox'][1] for b in kanan)
            tumpang = min(bawah_kiri, max(b['bbox'][3] for b in kanan)) - max(atas_kiri, atas_kanan)
            tinggi = max(bawah_kiri - atas_kiri, 1)
            penuh = [b for b in blok_teks if (b['bbox'][2] - b['bbox'][0]) > page.rect.width * K["full_width_ratio"]]
            dua_kolom = tumpang > tinggi * K["col_overlap_ratio"] and len(penuh) <= 2

        # --- Tandai watermark: teks miring (putar) atau raksasa-pudar.
        # Miring -> paragraf tengah (isi lestari). Raksasa-pudar yang
        # menindih isi -> watermark: masuk header Word abu-abu, lewati badan.
        def _lum_warna(c):
            try:
                c = int(c)
                return (0.299 * ((c >> 16) & 255) + 0.587 * ((c >> 8) & 255) + 0.114 * (c & 255)) / 255
            except Exception:
                return 0.0

        _wm_lewat = set()  # kunci bbox garis watermark -> lewati dari badan
        try:
            _semua_garis = []
            for _b in blok_teks:
                for _l in _b.get('lines', []):
                    _bb = _l.get('bbox', (0, 0, 0, 0))
                    _semua_garis.append(_bb)
            for _b in blok_teks:
                for _l in _b.get('lines', []):
                    _bb = _l.get('bbox', (0, 0, 0, 0))
                    _sps = [s for s in _l.get('spans', []) if (s.get('text') or '').strip()]
                    if not _sps:
                        continue
                    _mx = max(float(s.get('size', 11)) for s in _sps)
                    if _mx < 28:
                        continue
                    if not all(_lum_warna(s.get('color')) > 0.7 for s in _sps):
                        continue
                    _tindih = sum(1 for _o in _semua_garis
                                  if not (_o[2] <= _bb[0] or _o[0] >= _bb[2]
                                          or _o[3] <= _bb[1] or _o[1] >= _bb[3]))
                    # termasuk dirinya sendiri: >=2 berarti menindih >=1 baris isi
                    if _tindih >= 2:
                        _wm_lewat.add((round(_bb[0]), round(_bb[1]),
                                       round(_bb[2]), round(_bb[3])))
        except Exception:
            pass
        # --- Peluru vektor (kotak kecil tergambar, bukan glyph) & garis
        # aturan horizontal (underline heading/kontak ala CV).
        _peluru_v, _garis_bawah = set(), []
        try:
            _semua_baris = []
            for _b in blok_teks:
                for _l in _b.get('lines', []):
                    if any((s.get('text') or '').strip() for s in _l.get('spans', [])):
                        _semua_baris.append(_l)
            for _dr in page.get_drawings():
                _rr = _dr.get('rect')
                if _rr is None:
                    continue
                _dw, _dh = _rr.x1 - _rr.x0, _rr.y1 - _rr.y0
                _fl = _dr.get('fill')
                # kotak peluru: kecil, bujur, terisi gelap
                if _fl is not None and max(_fl) < 0.25 and 1.5 <= _dw <= 5.5 \
                        and 1.5 <= _dh <= 5.5 and abs(_dw - _dh) <= 1.2:
                    for _gl in _semua_baris:
                        _gb = _gl.get('bbox', (0, 0, 0, 0))
                        if abs((_rr.y0 + _rr.y1) / 2 - (_gb[1] + _gb[3]) / 2) <= 6.0 \
                                and _gb[0] - 30.0 <= _rr.x1 <= _gb[0] - 2.0:
                            _peluru_v.add(id(_gl))
                            break
                # garis aturan: horizontal panjang dan tipis
                _wd = _dr.get('width') or 0
                if _dh <= 4.0 and (_rr.x1 - _rr.x0) >= 150.0 \
                        and (_wd >= 0.4 or (_fl is not None and max(_fl) < 0.25)):
                    _garis_bawah.append((_rr.x0, (_rr.y0 + _rr.y1) / 2, _rr.x1))
            _gb2 = []
            for _g in sorted(_garis_bawah, key=lambda g: g[1]):
                if not _gb2 or _g[1] - _gb2[-1][1] > 2.0:
                    _gb2.append(_g)
            _garis_bawah = _gb2
        except Exception:
            _peluru_v, _garis_bawah = set(), []
        # Tabel borderless lintas-blok dideteksi dulu agar garisnya tak bocor
        # jadi paragraf; garis terpakai dilewati di loop baris nanti.
        try:
            _tbl_bl, _garis_pakai = _scan_tabel_halaman(blok_teks, body_size, page_cx, page.rect.width)
        except Exception:
            _tbl_bl, _garis_pakai = [], set()

        def _baris_dibuang(_l):
            """True bila baris pasti dilewati loop (watermark/putar/halaman)."""
            _bb = _l.get('bbox', (0, 0, 0, 0))
            if tuple(round(v) for v in _bb) in _wm_lewat:
                return True
            _ld = _l.get('dir', (1, 0)) or (1, 0)
            if len(_ld) > 1 and abs(_ld[1]) > 0.2:
                return True
            _tx = ' '.join((s.get('text') or '') for s in _l.get('spans', [])).strip()
            if not _tx:
                return True
            try:
                if (_bb[1] < page.rect.height * 0.12 or _bb[1] > page.rect.height * 0.88) \
                        and _NO_HAL.fullmatch(_tx):
                    return True
            except Exception:
                pass
            return False
        # Pasangan entri-tahun CV (kolom kanan tahun) -> tabel polos 2 kolom
        try:
            _tbl_th, _pakai_th = _pasang_tahun(blok_teks)
        except Exception:
            _tbl_th, _pakai_th = [], set()
        # Kotak tabel BERGARIS: teks di dalamnya milik tabel -> blok teks yang
        # bertumpuk diabaikan agar isi tidak ganda (paragraf + tabel).
        try:
            _tab_boxes = [tab.bbox for tab in page.find_tables()]
        except Exception:
            _tab_boxes = []

        def _di_dalam_tabel(bb):
            _cx, _cy = (bb[0] + bb[2]) / 2, (bb[1] + bb[3]) / 2
            return any(t[0] - 2 <= _cx <= t[2] + 2 and t[1] - 2 <= _cy <= t[3] + 2
                       for t in _tab_boxes)

        items = []  # (x0, y0, x1, y1, jenis, payload) -> diurut XY-cut
        for _tx0, _ty0, _tx1, _ty1, _rows in _tbl_bl:
            items.append((_tx0, _ty0, _tx1, _ty1, 'tabel', _rows))
        for _tx0, _ty0, _tx1, _ty1, _kind, _rows in _tbl_th:
            items.append((_tx0, _ty0, _tx1, _ty1, 'tabel_polos', _rows))
        for b in blok_teks:
            if kop_set and _norm_kop(' '.join(
                    s.get('text', '') for l in b.get('lines', [])
                    for s in l.get('spans', []))) in kop_set:
                continue  # kop berulang -> sudah di header Word
            _bb = b['bbox']
            items.append((_bb[0], _bb[1], _bb[2], _bb[3], 'blok', b))
        try:
            for tab in page.find_tables():
                r = tab.extract()
                if r and r[0]:
                    bb = tab.bbox
                    items.append((bb[0], bb[1], bb[2], bb[3], 'tabel', r))
        except Exception:
            pass
        try:
            rects_fn = getattr(page, 'get_image_rects', None)
            for img in page.get_images(full=True):
                xref = img[0]
                if rects_fn is not None:
                    try:
                        rs = rects_fn(xref)
                    except Exception:
                        rs = []
                    if rs:
                        for rc in rs:
                            items.append((rc[0], rc[1], rc[2], rc[3], 'gambar', (xref, rc)))
                        continue
                items.append((0, float('inf'), page.rect.width, float('inf'), 'gambar', (xref, None)))
        except Exception:
            pass
        items = _xy_urut(items, K)

        if dua_kolom:
            lebar_kolom_inci = min(6.0, lebar_konten_pt / 72 / 2)
        else:
            lebar_kolom_inci = min(6.0, lebar_konten_pt / 72)

        buf = []  # buffer paragraf: list[(spans, x0, lebar, y0, y1)]
        buf_tengah = False
        tunggu_sambung = False  # baris sebelumnya berakhir tanda hubung
        _prev_akhir = ""  # token terakhir baris sebelum (sambung tanpa hubung)
        _prev_penuh = False  # baris sebelum selebar konten (bukan akhir paragraf)
        _prev_bawah = 0.0  # y bawah baris sebelum (rapat = lanjutan paragraf)
        _gantung = False  # peluru Word terpisah baris -> tempel ke teks berikut
        _tunda = None  # panjang buf saat siram ditunda (paragraf nyambung)
        _peluru_pos = None  # (x0, y1, size) paragraf peluru vektor terakhir

        def siram_buffer():
            nonlocal buf, buf_tengah, tunggu_sambung, _prev_akhir, _prev_penuh, _prev_bawah, _gantung, _tunda, _peluru_pos
            if not buf:
                _prev_akhir, _prev_penuh, _prev_bawah = "", False, 0.0
                return
            _tunda = None  # flush aktual membatalkan penundaan
            _gantung = False  # paragraf baru memutus gantungan peluru
            p = doc.add_paragraph()
            # SATU BARIS: tempatkan PRESISI via indent (kebal substitusi font),
            # bukan alignment tengah/kanan yang bergeser mengikuti lebar teks.
            _satu = (len(buf) == 1)
            if buf_tengah and not _satu:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            elif len(buf) == 1 and not buf_tengah:
                # satu baris menempel kanan = rata kanan (mis. tanda tangan).
                # KECUALI lanjutan bungkus paragraf peluru vektor di kolom yang
                # sama -> tetap kiri + indent kolom (bukan rata kanan).
                _e = buf[0]
                _sz0 = float(_e[0][0][1].get('size', body_size)) if _e[0] else float(body_size)
                _lanjut = bool(_peluru_pos) and abs(_e[1] - _peluru_pos[0]) <= 4.0 \
                    and abs(_sz0 - _peluru_pos[2]) <= K["size_tol_pt"] \
                    and 0 <= (_e[3] - _peluru_pos[1]) <= _sz0 * 2.2
                if _lanjut:
                    _peluru_pos = (_e[1], _e[4], _sz0)
                else:
                    _peluru_pos = None
                if not _lanjut and _e[1] + _e[2] >= m_kiri_pt + lebar_konten_pt * K["single_right_edge"] \
                        and _e[1] > m_kiri_pt + lebar_konten_pt * K["single_right_x"]:
                    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                    try:
                        _ri = (m_kiri_pt + lebar_konten_pt) - (_e[1] + _e[2])
                        if _ri > 1:
                            p.paragraph_format.right_indent = Inches(_ri / 72)
                    except Exception:
                        pass
            else:
                # rata kanan? atau rata kiri-kanan (justify) seperti dokumen asli?
                _xs0 = [e[1] for e in buf]
                _ws = [e[2] for e in buf]
                _kn = [e[1] + e[2] for e in buf]
                if len(buf) >= 2 and (max(_kn) - min(_kn)) < K["right_tol_pt"] and (max(_xs0) - min(_xs0)) > K["right_gap_pt"]:
                    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                elif len(buf) >= 3 and all(w >= lebar_konten_pt * K["justify_minw_ratio"] for w in _ws[:-1]) \
                        and (max(_xs0) - min(_xs0)) < K["justify_tol_pt"]:
                    # justify = >=2 baris pembanding dengan tepi KIRI dan KANAN
                    # sejajar. Paragraf 2 baris terlalu ambigu -> aman jadi LEFT.
                    # (Uji regang-spasi dicabut: typesetting profesional
                    # me-justify via skala glyph, bukan celah kata.)
                    _knn = _kn[:-1]
                    _rata_kanan = (max(_knn) - min(_knn)) < K["justify_tol_pt"] if len(_knn) >= 2 \
                        else _ws[0] >= lebar_konten_pt * K["justify_single_full"]
                    if _rata_kanan:
                        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            # indent kiri mengikuti posisi teks asli; first-line bila baris 1 menjorok.
            # Untuk paragraf rata kanan/tengah multi-baris, posisi diatur
            # alignment -> indent nol. Baris TUNGGAL selalu diindent presisi
            # (tengah/kanan/kiri) agar x-nya sama persis dengan PDF.
            # Blok jauh menjorok yang menempel tepi kanan (tanda tangan) -> RIGHT.
            _x_first = buf[0][1]
            _x_kedua = buf[1][1] if len(buf) > 1 else _x_first
            _basis = min(_x_first, _x_kedua)
            _kanan_isi = max(e[1] + e[2] for e in buf)
            _rata_kanan_blok = (_basis > m_kiri_pt + lebar_konten_pt * 0.45) and \
                (_kanan_isi >= m_kiri_pt + lebar_konten_pt * 0.90)
            # Kolom teks sempit multi-baris sejajar kiri (mis. kolom kanan CV)
            # bukan teks rata kanan: kecualikan dari RIGHT, pakai indent kolom.
            _lebar_buf = max(e[1] + e[2] for e in buf) - min(e[1] for e in buf)
            _kolom_sempit = len(buf) >= 2 \
                and (max(e[1] for e in buf) - min(e[1] for e in buf)) <= 4.0 \
                and _lebar_buf < lebar_konten_pt * 0.5
            if _kolom_sempit:
                _rata_kanan_blok = False
            if p.alignment in (WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.CENTER):
                pass
            elif _satu or _rata_kanan_blok:
                if _rata_kanan_blok and not _satu:
                    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                else:
                    if _basis > m_kiri_pt + K["indent_tol_pt"]:
                        p.paragraph_format.left_indent = Inches((_basis - m_kiri_pt) / 72)
            else:
                if _basis > m_kiri_pt + K["indent_tol_pt"]:
                    p.paragraph_format.left_indent = Inches((_basis - m_kiri_pt) / 72)
                if _x_first > _x_kedua + K["firstline_min_pt"]:
                    p.paragraph_format.first_line_indent = Inches((_x_first - _x_kedua) / 72)
            # spasi baris mengikuti pitch garis asli (bukan default template)
            if len(buf) >= 2:
                try:
                    _pt = sorted(buf[i + 1][3] - buf[i][3] for i in range(len(buf) - 1))
                    _med_pt = _pt[(len(_pt) - 1) // 2]
                    _ls = min(2.5, max(1.0, round(_med_pt / max(float(body_size), 1), 2)))
                    p.paragraph_format.line_spacing = _ls
                except Exception:
                    pass
            for j, (spans, _blx, _blw, _by0, _by1) in enumerate(buf):
                for k, (teks, sp, _spasi) in enumerate(spans):
                    if j > 0 and k == 0:
                        p.add_run(' ')  # ganti baris visual -> selalu spasi
                    elif k > 0 and _spasi and not _lem_tanpa_spasi(spans[k - 1][0], teks):
                        p.add_run(' ')
                    tambah_run(p, teks, sp, body_size)
            if _garis_bawah:
                # garis aturan horizontal tepat di bawah blok -> border bawah
                # paragraf (underline heading/kontak ala CV)
                try:
                    _qx0 = min(e[1] for e in buf)
                    _qx1 = max(e[1] + e[2] for e in buf)
                    _qy1 = max(e[4] for e in buf)
                    for (_gx0, _gy, _gx1) in _garis_bawah:
                        if -4.0 <= _gy - _qy1 <= 14.0 and _gx0 < _qx1 and _gx1 > _qx0:
                            from docx.oxml import OxmlElement as _Ox2
                            _pPr = p._p.get_or_add_pPr()
                            _pBdr = _Ox2('w:pBdr')
                            _bt = _Ox2('w:bottom')
                            for _k, _v in (('w:val', 'single'), ('w:sz', '6'),
                                           ('w:space', '1'), ('w:color', '000000')):
                                _bt.set(qn(_k), _v)
                            _pBdr.append(_bt)
                            _pPr.append(_pBdr)
                            break
                except Exception:
                    pass
            buf = []
            buf_tengah = False
            tunggu_sambung = False
            _prev_akhir, _prev_penuh, _prev_bawah = "", False, 0.0

        _lebar_lokal = lebar_konten_pt
        for _ii, (_ix0, _iy0, _ix1, _iy1, jenis, payload) in enumerate(items):
            _buf_awal = len(buf)  # item yang tak menambah isi tak boleh flush
            if dua_kolom and (_ix1 - _ix0) < page.rect.width * 0.6:
                _lebar_lokal = lebar_konten_pt / 2
            else:
                _lebar_lokal = lebar_konten_pt
            if jenis == 'tabel':
                if _tunda is None or len(buf) != _tunda:
                    siram_buffer()
                tulis_tabel(doc, payload)
                continue
            if jenis == 'tabel_polos':
                if _tunda is None or len(buf) != _tunda:
                    siram_buffer()
                tulis_tabel_polos(doc, payload)
                continue
            if jenis == 'gambar':
                if _tunda is None or len(buf) != _tunda:
                    siram_buffer()
                xref, _rc = payload
                _lbr = (_ix1 - _ix0) / 72 if _ix1 > _ix0 else lebar_kolom_inci
                _asli = ((_rc[2] - _rc[0]) if _rc else 0) or None
                _tasli = ((_rc[3] - _rc[1]) if _rc else 0) or None
                tulis_gambar(doc, src, xref, min(max(_lbr, 1.0), lebar_kolom_inci),
                             lebar_asli_pt=_asli,
                             x_pt=(_rc[0] if _rc else None),
                             y_pt=(_rc[1] if _rc else None),
                             konten_pt=lebar_konten_pt,
                             tinggi_asli_pt=_tasli)
                # gambar kecil/hiasan: tanpa paragraf kosong tambahan
                continue
            b = payload
            if _blok_jadi_tabel(doc, b, body_size):
                continue  # blok kolom-sejajar -> sudah jadi tabel Word
            _baris = b.get('lines', [])
            for _li, l in enumerate(_baris):
                if id(l) in _garis_pakai or id(l) in _pakai_th:
                    _prev_akhir, _prev_penuh, _prev_bawah = "", False, 0.0
                    _gantung = False
                    continue  # garis milik tabel borderless/pasangan tahun -> lewati
                _lbx = l.get('bbox', (0, 0, 0, 0))
                if _di_dalam_tabel(_lbx):
                    _prev_akhir, _prev_penuh, _prev_bawah = "", False, 0.0
                    continue  # baris isi tabel bergaris -> sudah diwakili tabel Word
                _mentah = [s for s in l.get('spans', []) if s.get('text')]
                if not _mentah:
                    continue
                # Tandai pangkat atas/bawah: ukuran lebih kecil + baseline
                # menyimpang dari garis bawah baris (sitasi, rumus).
                _lb = l.get('bbox', (0, 0, 0, 0))
                _szs = sorted(s.get('size', body_size) for s in _mentah)
                _med = _szs[len(_szs) // 2]
                spans = []
                for s in _mentah:
                    _d = dict(s)
                    try:
                        _ssz = float(s.get('size', body_size))
                        _sb = s.get('bbox', _lb)
                        _naik = _lb[3] - _sb[3]
                        _turun = _sb[1] - _lb[1]
                        if _ssz < _med * 0.85 and _naik > 1.0:
                            _d['_sup'] = True
                        elif _ssz < _med * 0.85 and _turun > 1.0:
                            _d['_sub'] = True
                    except Exception:
                        pass
                    spans.append(_d)
                if not spans:
                    continue
                teks = ''.join(s['text'] for s in spans)
                # normalisasi spasi; JAGA tanda hubung akhir baris untuk digabung
                # (termasuk soft-hyphen U+00AD & hyphen unicode jurnal ilmiah)
                sambung = teks.rstrip().endswith(('-', '\u00ad', '\u2010', '\u2011')) \
                    and len(teks.strip()) > 1
                teks_bersih = ' '.join(teks.split())
                if not teks_bersih:
                    continue
                # buang karakter kontrol NULL/dll yang merusak XML Word
                teks_bersih = _bersih_xml(teks_bersih)
                if not teks_bersih:
                    continue
                teks_bersih = _perbaiki_spasi_hilang(teks_bersih)
                # Normalisasi ligatur & tanda kutip/strip tipografi agar teks
                # bisa dicari dan tidak pecah: ﬁ->fi, ’->', –->-, dst.
                for _lig, _asc in (('ﬁ', 'fi'), ('ﬂ', 'fl'), ('ﬀ', 'ff'),
                                   ('ﬃ', 'ffi'), ('ﬄ', 'ffl'),
                                   ('ﬆ', 'st'), ('ﬅ', 'ct'),
                                   ('æ', 'ae'), ('œ', 'oe'),
                                   ('‘', "'"), ('’', "'"), ('“', '"'), ('”', '"'),
                                   ('–', '-'), ('—', '-')):
                    teks_bersih = teks_bersih.replace(_lig, _asc)
                # Glyph tak-terpetakan font (U+FFFD): diapit huruf -> buang
                # (mis. organization�s), diapit spasi -> em-dash.
                teks_bersih = re.sub(r'(?<=[A-Za-z])\ufffd(?=[A-Za-z])', '', teks_bersih)
                teks_bersih = re.sub(r'\s+', ' ', teks_bersih.replace('\ufffd', ' — ')).strip()
                # lewati nomor halaman & kop halaman.
                # Nomor halaman terisolasi HANYA diabaikan bila di tepi atas /
                # bawah kertas: angka tahun di tengah isi ("2022") bukanlah
                # nomor halaman!
                _ly0 = l.get('bbox', (0, 0, 0, 0))[1]
                _di_tepi = _ly0 < page.rect.height * 0.12 or _ly0 > page.rect.height * 0.88
                if _di_tepi and _NO_HAL.fullmatch(teks_bersih):
                    continue
                tl = teks_bersih.lower()
                if len(teks_bersih) < K["kop_halaman_len"] and (tl.startswith('halaman ') or tl.startswith('page ')):
                    continue
                max_size = max(s.get('size', body_size) for s in spans)
                all_bold = all('bold' in (s.get('font', '').lower()) or (s.get('flags', 0) & 16) for s in spans)
                x0, y0, x1, y1 = l.get('bbox', (0, 0, 0, 0))
                # Watermark raksasa-pudar -> BUANG TOTAL dari hasil
                # (tidak ditulis ke badan maupun header).
                if (round(x0), round(y0), round(x1), round(y1)) in _wm_lewat:
                    _prev_akhir, _prev_penuh, _prev_bawah = "", False, 0.0
                    _gantung = False
                    continue
                # Teks miring/putar (cap/stempel/watermark): dilewati total
                # agar tidak ada sisa watermark pada hasil.
                _ldir = l.get('dir', (1, 0)) or (1, 0)
                if len(_ldir) > 1 and abs(_ldir[1]) > 0.2:
                    _prev_akhir, _prev_penuh, _prev_bawah = "", False, 0.0
                    _gantung = False
                    continue
                # Peluru Word yang terpisah baris dari teksnya (glyph Symbol
                # sendirian di barisnya) -> gantungkan ke baris teks berikut.
                if not _gantung and re.fullmatch(r'[\u2022\u00b7\ufffd\uf0b7\u25cf\u25aa]{1,2}', teks_bersih) \
                        and not _di_dalam_tabel(_lbx):
                    _gantung = True
                    _prev_akhir, _prev_penuh, _prev_bawah = "", False, 0.0
                    continue
                # tengah = baris yang midpoint-nya di tengah + celah simetris
                # (judul/kop). Bukan baris selebar penuh ( ambiguous ): batas
                # atas 0.93 konten. Bukan kandidat justify (itu multi-baris).
                _gap_kiri = x0 - m_kiri_pt
                _gap_kanan = (m_kiri_pt + lebar_konten_pt) - x1
                _lebar_rel = (x1 - x0) / max(lebar_konten_pt, 1)
                tengah = _lebar_rel < min(K["center_maxw_ratio"], 0.93) and \
                    abs((x0 + x1) / 2 - page_cx) < page.rect.width * K["center_tol_ratio"] and \
                    abs(_gap_kiri - _gap_kanan) < max(24.0, (x1 - x0) * 0.6)
                if tengah and (re.match(r'^[:;,a-z]', teks_bersih) or
                               re.fullmatch(r'\(?\d{1,3}[.)]?\s*', teks_bersih)):
                    # lanjutan kalimat (huruf kecil/':') & label angka polos
                    # bukan judul tengah -> paksa rata kiri
                    tengah = False
                jdl = (max_size >= body_size + K["head_d1"] and len(teks_bersih) < K["head_len1"] and (_lebar_rel < 0.88 or max_size >= body_size + K["head_lvl_gap"])) or (
                    all_bold and max_size >= body_size + K["head_d2"] and len(teks_bersih) < K["head_len2"])
                daftar = bool(re.match(r'^(\d+[.)]|[\u2022\-\*\u00b7\ufffd\uf0b7\u25cf\u25aa])\s+\S', teks_bersih))
                if _gantung and not jdl:
                    # tempel peluru gantungan ke baris ini lalu nilai ulang
                    teks_bersih = '• ' + teks_bersih
                    _gantung = False
                elif jdl:
                    _gantung = False  # judul: peluru nyasar dibuang
                daftar = bool(re.match(r'^(\d+[.)]|[\u2022\-\*\u00b7\ufffd\uf0b7\u25cf\u25aa])\s+\S', teks_bersih)) \
                    or id(l) in _peluru_v
                if jdl and not daftar and re.search(r'[A-Za-z,;:]$', teks_bersih):
                    # Calon judul yang baris BERIKUTNYA seformat & sejajar rapat
                    # adalah patahan paragraf badan berfont besar (bukan judul)
                    # -> turunkan jadi teks badan agar menggabung.
                    try:
                        _nl = _baris[_li + 1]
                        _nb = _nl.get('bbox', (0, 0, 0, 0))
                        _ns = [_s.get('size', body_size) for _s in _nl.get('spans', []) if _s.get('text')]
                        if _ns and abs(max(_ns) - max_size) <= K["size_tol_pt"] \
                                and abs(_nb[0] - x0) <= 4.0 \
                                and 0 <= (_nb[1] - y1) <= max_size * 1.5:
                            jdl = False
                    except (IndexError, TypeError):
                        pass
                if jdl or daftar:
                    siram_buffer()
                    _peluru_pos = None  # judul/daftar memutus rantai lanjutan peluru
                    if jdl:
                        level = 1 if max_size >= body_size + K["head_lvl_gap"] else (2 if max_size >= body_size + K["head_d1"] else 3)
                        # Judul ditulis dengan UKURAN/WARNA/FONT ASLI (bukan
                        # default style) agar sama persis dengan PDF sumber.
                        _hp = doc.add_paragraph(style=f'Heading {level}')
                        _px = None
                        for _s in spans:
                            _stx = _bersih_xml(' '.join(_s['text'].split()))
                            if not _stx:
                                continue
                            _sbb = _s.get('bbox', (x0, 0, x0, 0))
                            if _px is not None and _sbb[0] - _px > max(1.0, float(_s.get('size', 11)) * 0.25):
                                _ekt = _hp.text[-1:] if _hp.text else ''
                                if not _lem_tanpa_spasi(_ekt, _stx):
                                    _hp.add_run(' ')
                            _rh = tambah_run(_hp, _stx, _s, body_size)
                            try:
                                # tebal HARUS eksplisit agar tak mewarisi bold style Heading
                                _rh.font.bold = bool('bold' in (_s.get('font', '').lower()) or (_s.get('flags', 0) & 16))
                            except Exception:
                                pass
                            _px = _sbb[2]
                        if _garis_bawah:
                            # judul bergaris bawah di sumber -> border bawah
                            try:
                                for (_gx0, _gy, _gx1) in _garis_bawah:
                                    if -4.0 <= _gy - y1 <= 14.0 and _gx0 < x1 and _gx1 > x0:
                                        from docx.oxml import OxmlElement as _Ox3
                                        _hPr = _hp._p.get_or_add_pPr()
                                        _hBdr = _Ox3('w:pBdr')
                                        _hb = _Ox3('w:bottom')
                                        for _k, _v in (('w:val', 'single'), ('w:sz', '6'),
                                                       ('w:space', '1'), ('w:color', '000000')):
                                            _hb.set(qn(_k), _v)
                                        _hBdr.append(_hb)
                                        _hPr.append(_hBdr)
                                        break
                            except Exception:
                                pass
                    else:
                        if id(l) in _peluru_v:
                            # peluru vektor Word: Normal + '• ' + run berformat,
                            # menjorok kolom (gantung -10pt menyerupai sumber)
                            _pb = doc.add_paragraph()
                            _pb.style = doc.styles['Normal']
                            _awal = True
                            for _s in spans:
                                _stx = _bersih_xml(' '.join(_s['text'].split()))
                                if not _stx:
                                    continue
                                tambah_run(_pb, ('• ' if _awal else '') + _stx, _s, body_size)
                                _awal = False
                            if x0 > m_kiri_pt + K["indent_tol_pt"]:
                                _pb.paragraph_format.left_indent = Inches((x0 - m_kiri_pt) / 72)
                                _pb.paragraph_format.first_line_indent = Inches(-10 / 72)
                            _peluru_pos = (x0, y1, max_size)
                        else:
                            # peluru tak-terpetakan font (�, • Symbol) dinormalkan jadi • agar rapi
                            _dl = re.sub(r'^[\ufffd\u00b7\uf0b7\u25cf\u25aa]\s*', '• ', teks_bersih)
                            doc.add_paragraph(_dl,
                                style='List Bullet' if _dl[0] in '\u2022-*' else 'List Number')
                    continue
                # gabung ke paragraf mengalir bila sebaris wajar & sejajar.
                # Kecuali: ekor pendek yang simetris di bawah paragraf kiri
                # adalah baris TERAKHIR paragraf itu (bukan judul tengah).
                if buf and tengah and not buf_tengah and _prev_bawah > 0 \
                        and abs(max_size - buf[0][0][0][1].get('size', body_size)) <= K["size_tol_pt"] \
                        and 0 <= (y0 - _prev_bawah) <= max_size * 1.8:
                    tengah = False
                if buf and (tengah != buf_tengah or abs(max_size - buf[0][0][0][1].get('size', body_size)) > K["size_tol_pt"]):
                    siram_buffer()
                _w = (x1 - x0)
                # gabung span seformat yang bersebelahan agar run Word bersih.
                # Spasi antar-span mengikuti JARAK geometri: ada gap -> spasi,
                # menempel (ganti font tengah kata) -> tanpa spasi.
                _pot = []
                _prev_x1 = None
                _prev_trail = False
                for _s in spans:
                    _raw = _s['text']
                    _st = _bersih_xml(' '.join(_raw.split()))
                    if not _st:
                        continue
                    _sb = _s.get('bbox', (x0, 0, x0, 0))
                    _gap = (_sb[0] - _prev_x1) if _prev_x1 is not None else 999.0
                    _lead = len(_raw) != len(_raw.lstrip())
                    _need_sp = _prev_trail or _lead or \
                        _gap > max(1.0, float(_s.get('size', 11)) * 0.25)
                    _prev_x1 = _sb[2]
                    _prev_trail = len(_raw) != len(_raw.rstrip())
                    if _pot and _fmt_key(_pot[-1][1]) == _fmt_key(_s):
                        _spm = '' if _lem_tanpa_spasi(_pot[-1][0], _st) else (' ' if _need_sp else '')
                        _pot[-1] = (_pot[-1][0] + _spm + _st,
                                    _pot[-1][1], _pot[-1][2])
                    else:
                        _pot.append((_st, _s, _need_sp))
                if tunggu_sambung and buf and _pot:
                    # sambungan kata terpotong: tempel kata pertama ke ekor baris lalu
                    _t0, _s0, _sp0 = _pot[0]
                    _kata = _t0.split(' ', 1)
                    _spans, _blx, _blw, _bly0, _bly1 = buf.pop()
                    if _spans:
                        _lt, _ls, _lsp = _spans.pop()
                        _spans.append((_lt + _kata[0], _ls, _lsp))
                    if len(_kata) > 1 and _kata[1]:
                        _pot[0] = (_kata[1], _s0, True)
                    else:
                        _pot = _pot[1:]
                    buf.append((_spans, _blx, _blw, min(_bly0, y0), max(_bly1, y1)))
                    tunggu_sambung = False
                elif buf and not tunggu_sambung and _pot and _prev_akhir and _prev_penuh \
                        and re.search(r'[a-z]$', _prev_akhir) \
                        and re.match(r'^[a-z]', (_pot[0][0] or '')):
                    # sambung TANPA tanda hubung (maturit+ity, or+ganizations):
                    # gabung bila rangkaiannya kata Inggris valid. Baris penuh +
                    # batas huruf kecil membuat false-join sangat jarang.
                    _t0, _s0, _sp0 = _pot[0]
                    _w1 = _t0.split(' ', 1)[0]
                    _gab = (_prev_akhir + _w1).lower()
                    _ws_en = _kata_inggris()
                    _kandidat = {_gab}
                    if _gab.endswith('es'):
                        _kandidat.add(_gab[:-2])
                    if _gab.endswith('s'):
                        _kandidat.add(_gab[:-1])
                    if _ws_en and len(_gab) >= 5 and _kandidat & _ws_en:
                        _spans, _blx, _blw, _bly0, _bly1 = buf.pop()
                        if _spans:
                            _lt, _ls, _lsp = _spans.pop()
                            _spans.append((_lt + _w1, _ls, _lsp))
                        _sisa = _t0[len(_w1):].lstrip()
                        _pot = ([(_sisa, _s0, True)] + _pot[1:]) if _sisa else _pot[1:]
                        buf.append((_spans, _blx, _blw, min(_bly0, y0), max(_bly1, y1)))
                if not buf:
                    buf_tengah = bool(tengah and len(teks_bersih) < 150)
                if sambung:
                    # "pemantau-an" + "lanjutan" -> "pemantauan": tanpa spasi
                    if _pot:
                        _t, _s, _sp = _pot[-1]
                        while len(_t) > 1 and _t[-1] in ("-", "\u00ad", "\u2010", "\u2011"):
                            _t = _t[:-1]
                        _pot[-1] = (_t, _s, _sp)
                    # (tanda hubung: ASCII '-', U+00AD, U+2010, U+2011)
                    if _pot:
                        buf.append((_pot, x0, _w, y0, y1))
                    tunggu_sambung = True
                elif _pot:
                    # satu baris visual = satu entri agar indent & justify akurat
                    buf.append((_pot, x0, _w, y0, y1))
                _tokas = teks_bersih.split()
                _prev_akhir = _tokas[-1] if _tokas else ""
                _prev_penuh = _w >= _lebar_lokal * 0.88
                _prev_bawah = y1
            # akhir blok: jeda paragraf (antar blok beda topik) — KECUALI
            # baris berikut jelas lanjutan kalimat (patahan paragraf Word):
            # buffer dipertahankan agar menggabung, bukan pecah per baris.
            _mau_siram = True
            if buf and _ii + 1 < len(items) and len(buf) != _buf_awal:
                # cari blok teks berikut yang BERISI (lewati blok yang
                # seluruh barisnya dibuang: watermark, garis tabel, dsb.)
                # gambar yang seluruhnya di atas celah juga dilewati.
                _nxt = None
                for _jj in range(_ii + 1, len(items)):
                    _jk = items[_jj][4]
                    if _jk == 'blok':
                        _cb = items[_jj][5]
                        _cl = [l for l in _cb.get('lines', [])
                               if any((s.get('text') or '').strip() for s in l.get('spans', []))]
                        if _cl and all(_baris_dibuang(l) for l in _cl):
                            continue
                        _nxt = _cb
                        break
                    if _jk == 'gambar' and items[_jj][3] <= _prev_bawah + 1.0:
                        continue
                    break
                if _nxt is not None:
                    _nls = [l for l in _nxt.get('lines', [])
                            if any((s.get('text') or '').strip() for s in l.get('spans', []))]
                    if _nls:
                        _nl0 = _nls[0]
                        _nb0 = _nl0.get('bbox', (0, 0, 0, 0))
                        _nsp0 = [s for s in _nl0.get('spans', []) if (s.get('text') or '').strip()]
                        _ntx = ' '.join(s.get('text', '') for s in _nsp0).strip()
                        _nsize = max(float(s.get('size', body_size)) for s in _nsp0)
                        _bl = None
                        for _ebb in reversed(buf):
                            if _ebb[0]:
                                _bl = _ebb
                                break
                        if _bl is not None:
                            _bsize = float(_bl[0][0][1].get('size', body_size))
                            _term = bool(re.search(r'[.!?…:;]$', _prev_akhir or ''))
                            _daf = bool(re.match(r'^(\d+[.)]|[\u2022\-\*\u00b7\ufffd\uf0b7\u25cf\u25aa])\s*\S', _ntx)) \
                                or id(_nl0) in _peluru_v
                            if (not _term and not _daf
                                    and abs(_nsize - _bsize) <= K["size_tol_pt"]
                                    and abs(_nb0[0] - _bl[1]) <= 4.0
                                    and 0 <= (_nb0[1] - _prev_bawah) <= _nsize * 1.5
                                    and (_prev_penuh or re.match(r'^[a-z]', _ntx or ''))):
                                _mau_siram = False
            _npara = len(doc.paragraphs)
            if buf and len(buf) == _buf_awal:
                pass  # item tak menambah baris (terlewati semua): tahan buffer
            elif _mau_siram:
                siram_buffer()
            else:
                _tunda = len(buf)  # paragraf berlanjut ke item berikut
            if _mau_siram and len(doc.paragraphs) > _npara and _ii + 1 < len(items):
                try:
                    _gap = items[_ii + 1][1] - b['bbox'][3]
                    _sa = min(24.0, max(0.0, (_gap - body_size * 0.6) * 0.7))
                    doc.paragraphs[-1].paragraph_format.space_after = Pt(_sa)
                except Exception:
                    pass
        siram_buffer()

    # Buang paragraf kosong di ujung dokumen (sisa pemisah tabel/gambar)
    # agar tidak menambah halaman kosong / menggeser tata letak.
    # Sama untuk paragraf kosong bawaan template di baris pertama.
    # Paragraf yang HANYA berisi gambar melayang (anchor) digabung run-nya
    # ke paragraf teks pertama agar tak menyisakan baris kosong.
    try:
        while len(doc.paragraphs) > 1 and not doc.paragraphs[-1].text.strip() \
                and not doc.paragraphs[-1].runs:
            _p_akhir = doc.paragraphs[-1]._p
            _p_akhir.getparent().remove(_p_akhir)
        while len(doc.paragraphs) > 1 and not doc.paragraphs[0].text.strip() \
                and not doc.paragraphs[0].runs:
            _p_awal = doc.paragraphs[0]._p
            _p_awal.getparent().remove(_p_awal)
        _tuan = next((pp for pp in doc.paragraphs if pp.text.strip()), None)
        for _pp in list(doc.paragraphs):
            if not _pp.text.strip() and _pp._p.xpath('.//wp:anchor') \
                    and not _pp._p.xpath('.//wp:inline'):
                if _tuan is None or _pp is _tuan:
                    for _r in _pp.runs:
                        try:
                            _r.font.size = Pt(1)
                        except Exception:
                            pass
                    _pp.paragraph_format.space_after = Pt(0)
                    _pp.paragraph_format.space_before = Pt(0)
                    if _tuan is None:
                        _tuan = _pp
                else:
                    for _r in list(_pp.runs):
                        _tuan._p.append(_r._r)
                    _pp._p.getparent().remove(_pp._p)
    except Exception:
        pass

    # Watermark DINONAKTIFKAN (permintaan pemilik): teks watermark yang
    # terdeteksi tetap dilewati dari badan (lihat _wm_lewat di loop baris),
    # tetapi TIDAK direplikasi ke mana pun — hasil bersih tanpa watermark.
    doc.core_properties.title = os.path.basename(pdf_path)
    doc.save(docx_path)
    src.close()


def _pdf2word_flowing(pdf_path, docx_path):
    """Mode 'Teks Mengalir': bangun ulang dokumen bersih yang mudah diedit.

    - Paragraf digabung per blok (bukan per baris visual) agar teks bisa reflow.
    - Judul terdeteksi dari ukuran font dominan; daftar & rata tengah dipertahankan.
    - Tabel asli (.find_tables) ditulis ulang jadi tabel Word; gambar disematkan.
    """
    import fitz
    from docx import Document
    from docx.shared import Pt, Inches
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml.ns import qn
    import re

    src = fitz.open(pdf_path)
    doc = Document()

    # Gaya dasar rapi
    normal = doc.styles['Normal']
    normal.font.name = 'Calibri'
    normal.font.size = Pt(11)
    normal.paragraph_format.space_after = Pt(6)
    normal.paragraph_format.line_spacing = 1.15
    for level, size in ((1, 16), (2, 14), (3, 12)):
        hs = doc.styles[f'Heading {level}']
        hs.font.name = 'Calibri'
        hs.font.size = Pt(size)
        hs.font.bold = True

    def shade_cell(cell, hex_color):
        tcPr = cell._tc.get_or_add_tcPr()
        shd = tcPr.makeelement(qn('w:shd'), {
            qn('w:val'): 'clear', qn('w:color'): 'auto', qn('w:fill'): hex_color})
        tcPr.append(shd)

    for pno, page in enumerate(src):
        if pno > 0:
            doc.add_page_break()
        data = page.get_text('dict')
        # Ukuran font tubuh = modus ukuran baris
        sizes = []
        for b in data.get('blocks', []):
            for l in b.get('lines', []):
                for s in l.get('spans', []):
                    if s.get('text', '').strip():
                        sizes.append(round(s.get('size', 11)))
        body_size = max(set(sizes), key=sizes.count) if sizes else 11
        page_cx = page.rect.width / 2

        for b in data.get('blocks', []):
            if b.get('type', 0) != 0:
                continue
            for l in b.get('lines', []):
                spans = [s for s in l.get('spans', []) if s.get('text')]
                if not spans:
                    continue
                text = ''.join(s['text'] for s in spans).strip()
                if not text:
                    continue
                max_size = max(s.get('size', body_size) for s in spans)
                all_bold = all('bold' in (s.get('font', '').lower()) or (s.get('flags', 0) & 16) for s in spans)
                x0, _, x1, _ = l.get('bbox', (0, 0, 0, 0))
                centered = abs((x0 + x1) / 2 - page_cx) < page.rect.width * 0.05

                # Judul: font jauh lebih besar / tebal + pendek
                if (max_size >= body_size + 3 and len(text) < 120) or (all_bold and max_size >= body_size + 1.5 and len(text) < 80):
                    level = 1 if max_size >= body_size + 5 else (2 if max_size >= body_size + 3 else 3)
                    doc.add_heading(text, level=level)
                    continue
                # Daftar bernomor / berpoin
                if re.match(r'^(\d+[.)]|[\u2022\-\*])\s+\S', text):
                    doc.add_paragraph(text, style='List Bullet' if text[0] in '\u2022-*' else 'List Number')
                    continue
                p = doc.add_paragraph()
                run = p.add_run(text)
                run.bold = all_bold
                if max_size > body_size + 0.5:
                    run.font.size = Pt(min(max_size, 18))
                if centered and len(text) < 150:
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER

        # Tabel asli -> tabel Word bergaya
        try:
            for tab in page.find_tables():
                rows = tab.extract()
                if not rows:
                    continue
                t = doc.add_table(rows=len(rows), cols=len(rows[0]))
                t.style = 'Table Grid'
                for ri, row in enumerate(rows):
                    for ci, val in enumerate(row):
                        cell = t.cell(ri, ci)
                        cell.text = '' if val is None else str(val).strip()
                        for par in cell.paragraphs:
                            par.style = doc.styles['Normal']
                        if ri == 0:
                            for par in cell.paragraphs:
                                for r in par.runs:
                                    r.bold = True
                            shade_cell(cell, 'D9E2F3')
                doc.add_paragraph()
        except Exception:
            pass

        # Gambar disematkan (maks lebar 6 inci, deduplikasi per xref)
        # CATATAN Windows: file temp harus DITUTUP dulu sebelum pix.save menulis
        # ke path yang sama, kalau tidak PermissionError dan gambar hilang diam-diam.
        import tempfile as _tf
        seen = set()
        for img in page.get_images(full=True):
            xref = img[0]
            if xref in seen:
                continue
            seen.add(xref)
            tmp_path = None
            try:
                pix = fitz.Pixmap(src, xref)
                if pix.n - pix.alpha > 3:
                    pix = fitz.Pixmap(fitz.csRGB, pix)
                fd, tmp_path = _tf.mkstemp(suffix='.png')
                os.close(fd)
                pix.save(tmp_path)
                doc.add_picture(tmp_path, width=Inches(6))
            except Exception:
                continue
            finally:
                try:
                    if tmp_path and os.path.exists(tmp_path):
                        os.unlink(tmp_path)
                except OSError:
                    pass

    doc.core_properties.title = os.path.basename(pdf_path)
    doc.save(docx_path)
    src.close()


@app.route('/api/pdf-to-word', methods=['POST'])
def pdf_to_word():
    import uuid
    if 'file' not in request.files:
        return jsonify({"error": "Tidak ada file yang diunggah."}), 400

    file = request.files['file']
    if not file.filename:
        return jsonify({"error": "Tidak ada file yang dipilih."}), 400

    filename = secure_filename(file.filename)
    if os.path.splitext(filename)[1].lower() != '.pdf':
        return jsonify({"error": "File harus berformat PDF (.pdf)."}), 400

    mode = (request.form.get('mode') or 'layout').strip().lower()
    if mode not in ('layout', 'flowing'):
        mode = 'layout'

    tag = uuid.uuid4().hex[:8]
    base = filename.rsplit('.', 1)[0]
    input_path = os.path.join(UPLOAD_FOLDER, f"{tag}_{filename}")
    output_filename = base + '.docx'
    output_path = os.path.join(OUTPUT_FOLDER, f"{tag}_{output_filename}")

    try:
        file.save(input_path)

        # Validasi: ukuran, sihir PDF, enkripsi, jumlah halaman
        if os.path.getsize(input_path) > PDF2WORD_MAX_MB * 1024 * 1024:
            return jsonify({"error": f"Ukuran PDF maksimal {PDF2WORD_MAX_MB} MB."}), 400
        with open(input_path, 'rb') as fh:
            if fh.read(5) != b'%PDF-':
                return jsonify({"error": "File rusak atau bukan PDF valid."}), 400

        import fitz
        try:
            probe = fitz.open(input_path)
        except Exception:
            return jsonify({"error": "File PDF rusak atau tidak valid. Coba buka di penampil PDF lalu simpan ulang."}), 400
        try:
            if probe.needs_pass:
                return jsonify({"error": "PDF terkunci kata sandi. Buka kunci dulu lalu unggah ulang."}), 400
            n_pages = probe.page_count
            if n_pages < 1:
                return jsonify({"error": "PDF tidak berisi halaman."}), 400
            if n_pages > PDF2WORD_MAX_PAGES:
                return jsonify({"error": f"PDF terlalu panjang (maks {PDF2WORD_MAX_PAGES} halaman)."}), 400
            sample = ''.join(probe[i].get_text() for i in range(min(n_pages, 3)))
            is_scan = len(sample.strip()) < 50
        finally:
            probe.close()

        warning = None
        if is_scan:
            # PDF pindaian/gambar: baca tulisannya dengan OCR agar hasilnya
            # berupa TEKS Word yang bisa diedit (bukan sekadar gambar).
            try:
                _pdf2word_ocr(input_path, output_path,
                              dpi=int(_muat_config_konversi().get("ocr_dpi", 150)))
                mode = 'ocr'
                warning = (f"PDF pindaian {n_pages} halaman dibaca otomatis (OCR). "
                           "Teks sudah bisa diedit — mohon periksa ejaan sebelum dipakai resmi.")
            except Exception as ocr_err:
                print(f"[PDF2Word] OCR gagal ({ocr_err}), fallback ke mode {mode}.")
                warning = ("OCR tidak tersedia saat ini, hasil berisi gambar halaman. "
                           "Coba lagi nanti untuk teks yang bisa diedit.")
                if mode == 'flowing':
                    _pdf2word_flowing_v2(input_path, output_path)
                else:
                    cv = Converter(input_path)
                    try:
                        cv.convert(
                            output_path, start=0, end=None,
                            multi_processing=(n_pages >= 4),
                            delete_end_line_hyphen=True,
                            connected_border_tolerance=1.0,
                            min_section_height=15.0,
                        )
                    finally:
                        cv.close()
        elif mode == 'flowing':
            _pdf2word_flowing_v2(input_path, output_path)
        else:
            # Mode Layout Akurat: pdf2docx yang disetel untuk kerapian maksimal
            # - multi_processing: jauh lebih cepat di PDF banyak halaman
            # - delete_end_line_hyphen: sambung kata terpotong tanda hubung
            # - connected_border_tolerance: sambungan garis tabel lebih utuh
            # - min_section_height: bagian kecil (kop/header) tidak hilang
            cv = Converter(input_path)
            try:
                cv.convert(
                    output_path, start=0, end=None,
                    multi_processing=(n_pages >= 4),
                    delete_end_line_hyphen=True,
                    connected_border_tolerance=1.0,
                    min_section_height=15.0,
                )
            finally:
                cv.close()

        # Metadata dokumen (+ buang watermark bila mode layout)
        try:
            from docx import Document as _Doc
            _d = _Doc(output_path)
            if mode == 'layout':
                # pdf2docx menyalin mentah termasuk watermark: bersihkan
                try:
                    _n_bersih = _buang_watermark_docx(output_path, _daftar_teks_watermark(input_path))
                    if _n_bersih:
                        print(f"[PDF2Word] { _n_bersih} paragraf watermark dibuang (mode layout).")
                except Exception as _e_wm:
                    print(f"[PDF2Word] gagal buang watermark: {_e_wm}")
                _d = _Doc(output_path)
            _d.core_properties.title = base
            _d.core_properties.comments = f"Dikonversi dari {filename} (mode {mode}) via K.O.M.A.N.D.O"
            _d.save(output_path)
        except Exception:
            pass

        resp = send_file(output_path, as_attachment=True, download_name=output_filename)
        resp.headers['X-PDF2Word-Pages'] = str(n_pages)
        resp.headers['X-PDF2Word-Mode'] = mode
        resp.headers['Access-Control-Expose-Headers'] = 'X-PDF2Word-Warning, X-PDF2Word-Pages, X-PDF2Word-Mode'
        if warning:
            # HTTP header wajib latin-1: ubah —/“/” dsb jadi padanan ASCII
            aman = (warning.replace('—', '-').replace('–', '-')
                    .replace('“', '"').replace('”', '"').replace('‘', "'").replace('’', "'"))
            aman = aman.encode('latin-1', 'replace').decode('latin-1')
            resp.headers['X-PDF2Word-Warning'] = aman
        return resp
    except Exception as e:
        import traceback
        traceback.print_exc()
        try:
            if os.path.exists(output_path):
                os.remove(output_path)
        except OSError:
            pass
        return jsonify({"error": f"Gagal mengonversi PDF: {str(e)}"}), 500
    finally:
        try:
            if os.path.exists(input_path):
                os.remove(input_path)
        except OSError:
            pass

# API 2: Konversi Word ke PDF (MS Word asli -> format 100% sama)
WORD2PDF_MAX_MB = 50
import threading as _threading
_WORD2PDF_LOCK = _threading.Lock()  # Word COM tidak tahan request bersamaan
_SCAN_LOCK = _threading.Lock()  # scanner fisik juga antre satu per satu


def _wia_devices():
    """Daftar perangkat WIA (scanner). Kembalikan list dict."""
    import win32com.client
    dm = win32com.client.Dispatch('WIA.DeviceManager')
    hasil = []
    for i in range(1, dm.DeviceInfos.Count + 1):
        try:
            info = dm.DeviceInfos(i)
            props = {}
            for nama in ('Name', 'Description', 'Port'):
                try:
                    props[nama] = info.Properties(nama).Value
                except Exception:
                    props[nama] = ''
            hasil.append({'index': i, 'id': info.DeviceID,
                          'nama': props.get('Name') or '',
                          'deskripsi': props.get('Description') or ''})
        except Exception:
            continue
    return hasil


@app.route('/api/scan-epson/status', methods=['GET'])
def scan_epson_status():
    """Status scanner fisik via WIA — TANPA Node/npm. Dipakai tombol Scan.
    Sengaja TANPA login agar dropdown driver tetap terisi; scan-nya
    sendiri (/api/scan-epson POST) cukup wajib login."""
    try:
        import pythoncom as _pc
        _pc.CoInitialize()
    except Exception:
        pass
    try:
        devs = _wia_devices()
        return jsonify({'online': bool(devs), 'jumlah': len(devs), 'perangkat': devs})
    except Exception as e:
        return jsonify({'online': False, 'error': f'WIA tidak tersedia: {e}'}), 200
    finally:
        try:
            import pythoncom as _pc2
            _pc2.CoUninitialize()
        except Exception:
            pass


@app.route('/api/scan-epson', methods=['POST'])
@login_required
def scan_epson():
    """Pindai langsung via WIA (tanpa Node). Kembalikan PDF multi-halaman.
    Mendukung flatbed (mis. Epson L3210) maupun ADF (mis. Epson DS-410,
    termasuk duplex bila diminta).
    Body JSON opsional: {dpi, colorMode(color|grayscale|bw), duplex, maxPages, prefix, deviceIndex, deviceId}."""
    import tempfile
    import uuid
    data = request.get_json(silent=True) or {}
    try:
        dpi = int(data.get('dpi') or 200)
    except Exception:
        dpi = 200
    dpi = min(max(dpi, 75), 600)
    mode_warna = (data.get('colorMode') or 'color').strip().lower()
    duplex = bool(data.get('duplex'))
    try:
        maks = int(data.get('maxPages') or 0)
    except Exception:
        maks = 0
    maks = maks if 0 < maks <= 50 else 50
    prefix = ''.join(c for c in str(data.get('prefix') or 'disposisi') if c.isalnum() or c in ('-', '_'))[:40] or 'disposisi'
    try:
        import pythoncom as _pc
        _pc.CoInitialize()
    except Exception:
        pass
    tmpdir = tempfile.mkdtemp(prefix='scan_epson_')
    pngs = []
    try:
        with _SCAN_LOCK:
            import win32com.client
            dm = win32com.client.Dispatch('WIA.DeviceManager')
            if dm.DeviceInfos.Count < 1:
                return jsonify({"success": False, "error": "Tidak ada scanner terdeteksi di server (WIA). Jika scanner USB terpasang di komputer Anda, jalankan Scan Agent lokal lalu coba lagi. Jika scanner terpasang di server, sambungkan melalui USB atau WiFi dan pastikan driver terpasang.", "code": "NO_SCANNER"}), 404
            # Pilih perangkat: index / deviceId dari request, default perangkat pertama.
            # Berlaku untuk USB maupun jaringan (WiFi) — semua tampil di WIA.
            dev = None
            _minta_idx = None
            try:
                _minta_idx = int(data.get('deviceIndex') or 0)
            except Exception:
                _minta_idx = 0
            _minta_id = str(data.get('deviceId') or '').strip()
            if _minta_id:
                for _i in range(1, dm.DeviceInfos.Count + 1):
                    try:
                        if dm.DeviceInfos(_i).DeviceID == _minta_id:
                            dev = dm.DeviceInfos(_i).Connect()
                            break
                    except Exception:
                        continue
                if dev is None:
                    return jsonify({"success": False, "error": "Perangkat yang dipilih tidak ditemukan."}), 404
            else:
                _idx = _minta_idx if 1 <= _minta_idx <= dm.DeviceInfos.Count else 1
                dev = dm.DeviceInfos(_idx).Connect()
            # ADF feeder (+duplex bila diminta); flatbed-only tetap jalan sekali
            _ada_feeder = False
            try:
                dev.Properties('Document Handling Select').Value = 1 | (4 if duplex else 0)
                _ada_feeder = True
            except Exception:
                pass
            item = dev.Items(1)
            try:
                item.Properties('Horizontal Resolution').Value = dpi
                item.Properties('Vertical Resolution').Value = dpi
            except Exception:
                pass
            try:
                # Wajib 1 halaman per transfer (penting untuk ADF seperti DS-410
                # agar loop multi-halaman menarik lembar satu per satu).
                item.Properties('Pages').Value = 1
            except Exception:
                pass
            try:
                item.Properties('Current Intent').Value = {'color': 1, 'grayscale': 2, 'bw': 4}.get(mode_warna, 1)
            except Exception:
                pass

            def _feeder_siap():
                try:
                    return bool(dev.Properties('Document Handling Status').Value & 1)
                except Exception:
                    return True

            # Epson L3210 = flatbed (tanpa ADF). Sebagian driver Epson/WIA
            # hanya mau transfer bila format PNG disebut eksplisit, jadi
            # sediakan fallback CommonDialog + FormatID PNG.
            _FMT_PNG = '{B96B3CAF-0728-11D3-9D7B-0000F81EF32E}'

            def _transfer_png():
                try:
                    return item.Transfer()
                except Exception:
                    pass
                try:
                    dlg = win32com.client.Dispatch('WIA.CommonDialog')
                    return dlg.ShowTransfer(item, _FMT_PNG, False)
                except Exception as e2:
                    raise e2

            sumber = 'ADF' if _ada_feeder else 'kaca scanner (flatbed)'
            n = 0
            while n < maks:
                if n > 0 and not _feeder_siap():
                    break
                if n >= 1 and not _ada_feeder:
                    break  # flatbed tanpa ADF: cukup satu halaman
                try:
                    img = _transfer_png()
                except Exception as e:
                    if n == 0:
                        return jsonify({"success": False, "error": f"Gagal memindai dari {sumber} (belum ada dokumen?): {e}", "code": "TRANSFER_GAGAL"}), 500
                    break
                fp = os.path.join(tmpdir, f'p{n + 1}.png')
                try:
                    img.SaveFile(fp)
                except Exception:
                    break
                pngs.append(fp)
                n += 1
                if n == 1 and not _feeder_siap():
                    # sekali jalan bila feeder langsung kosong sesudahnya
                    pass
            if not pngs:
                return jsonify({"success": False, "error": f"Tidak ada halaman terpindai. Letakkan dokumen di {sumber} lalu coba lagi.", "code": "PAPER_EMPTY"}), 400
            import fitz
            pdf = fitz.open()
            for fp in pngs:
                # WIA kadang menyimpan BMP walau ekstensi .png — Pixmap
                # membaca isi file (bukan ekstensi). Halaman ditempel
                # sebagai JPEG terkompresi agar PDF kecil (scan warna A4
                # 200dpi mentah bisa >10MB dan ditolak lampiran).
                pix = fitz.Pixmap(fp)
                try:
                    if pix.alpha:
                        pix = fitz.Pixmap(pix, 0)
                    if pix.n > 3:
                        try:
                            pix = fitz.Pixmap(fitz.csRGB, pix)
                        except Exception:
                            pass
                    w_pt = pix.width * 72.0 / dpi
                    h_pt = pix.height * 72.0 / dpi
                    page = pdf.new_page(width=w_pt, height=h_pt)
                    page.insert_image(page.rect, stream=pix.tobytes('jpg', jpg_quality=80))
                finally:
                    try:
                        pix = None
                    except Exception:
                        pass
            nama = f"{prefix}_{uuid.uuid4().hex[:8]}.pdf"
            out = os.path.join(tmpdir, nama)
            pdf.save(out)
            pdf.close()
            with open(out, 'rb') as fh:
                data_pdf = fh.read()
            resp = app.response_class(data_pdf, mimetype='application/pdf')
            resp.headers['Content-Disposition'] = f'attachment; filename="{nama}"'
            resp.headers['X-Scan-Pages'] = str(len(pngs))
            resp.headers['X-Scan-Source'] = sumber
            resp.headers['Access-Control-Expose-Headers'] = 'X-Scan-Pages, X-Scan-Source'
            return resp
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({"success": False, "error": f"Gagal memindai: {e}"}), 500
    finally:
        for fp in pngs:
            try:
                if os.path.exists(fp):
                    os.remove(fp)
            except OSError:
                pass
        try:
            import pythoncom as _pc2
            _pc2.CoUninitialize()
        except Exception:
            pass

# ===================== EDIT PDF LANGSUNG (ketik di atas PDF) =====================
# Frontend me-render PDF via pdf.js lalu menumpuk kotak teks yang bisa diketik
# tepat di atas posisi kata asli (koordinat dari endpoint parse). Saat simpan,
# backend menutup teks lama (redact putih) lalu menulis teks baru di koordinat sama.
@app.route('/api/pdf-edit/parse', methods=['POST'])
def pdf_edit_parse():
    import uuid
    if 'file' not in request.files:
        return jsonify({"error": "Tidak ada file yang diunggah."}), 400
    file = request.files['file']
    if not file.filename or os.path.splitext(file.filename)[1].lower() != '.pdf':
        return jsonify({"error": "File harus berformat PDF (.pdf)."}), 400
    tag = uuid.uuid4().hex[:12]
    simpan = os.path.join(UPLOAD_FOLDER, f"edit_{tag}.pdf")
    try:
        file.save(simpan)
        import fitz
        doc = fitz.open(simpan)
        try:
            if doc.needs_pass:
                return jsonify({"error": "PDF terkunci kata sandi."}), 400
            if doc.page_count > 30:
                return jsonify({"error": "Maksimal 30 halaman untuk edit langsung."}), 400
            pages = []
            for pi in range(doc.page_count):
                pg = doc[pi]
                r = pg.rect
                teks = (pg.get_text() or '').strip()
                words = pg.get_text('words') or []
                # Kelompokkan kata per baris (block, line) agar jadi satu kotak edit.
                garis = {}
                for w in words:
                    x0, y0, x1, y1, kata = w[0], w[1], w[2], w[3], w[4]
                    key = (w[5], w[6])
                    g = garis.setdefault(key, {'x0': x0, 'y0': y0, 'x1': x1, 'y1': y1, 'kata': []})
                    g['x0'] = min(g['x0'], x0); g['y0'] = min(g['y0'], y0)
                    g['x1'] = max(g['x1'], x1); g['y1'] = max(g['y1'], y1)
                    g['kata'].append((x0, kata))
                lines = []
                for g in garis.values():
                    g['kata'].sort(key=lambda t: t[0])
                    lines.append({'x0': round(g['x0'], 1), 'y0': round(g['y0'], 1),
                                  'x1': round(g['x1'], 1), 'y1': round(g['y1'], 1),
                                  'text': ' '.join(k for _, k in g['kata'])})
                lines.sort(key=lambda l: (round(l['y0'] / 10), l['x0']))
                pages.append({'width': round(float(r.width), 1), 'height': round(float(r.height), 1),
                              'is_scan': len(teks) < 20, 'lines': lines[:400]})
            return jsonify({"file_token": tag, "pages": pages}), 200
        finally:
            doc.close()
    except Exception as e:
        try:
            if os.path.exists(simpan):
                os.remove(simpan)
        except OSError:
            pass
        return jsonify({"error": f"Gagal membaca PDF: {e}"}), 500


@app.route('/api/pdf-edit/save', methods=['POST'])
def pdf_edit_save():
    import uuid
    data = request.get_json(force=True, silent=True) or {}
    token = (data.get('file_token') or '').strip()
    edits = data.get('edits') or []
    if not token or not edits:
        return jsonify({"error": "Tidak ada perubahan untuk disimpan."}), 400
    asal = os.path.join(UPLOAD_FOLDER, f"edit_{token}.pdf")
    if not os.path.exists(asal):
        return jsonify({"error": "Berkas asli kedaluwarsa. Unggah ulang PDF-nya."}), 404
    try:
        import fitz
        doc = fitz.open(asal)
        try:
            for e in edits[:2000]:
                try:
                    pi = int(e.get('page', 0))
                    if pi < 0 or pi >= doc.page_count:
                        continue
                    pg = doc[pi]
                    teks_baru = str(e.get('text') or '')
                    dihapus = bool(e.get('hapus'))
                    baru = bool(e.get('is_new'))
                    # Format per baris dari frontend (fallback = lama).
                    try:
                        fs = float(e.get('fontsize') or 0)
                    except (TypeError, ValueError):
                        fs = 0
                    bold = bool(e.get('bold'))
                    italic = bool(e.get('italic'))
                    warna = e.get('color') or ''
                    if baru:
                        rect = fitz.Rect(float(e['x0']), float(e['y0']),
                                         float(e['x0']) + 200, float(e['y0']) + max(10.0, fs or 12))
                    else:
                        rect = fitz.Rect(float(e['x0']), float(e['y0']), float(e['x1']), float(e['y1']))
                    if not baru:
                        # Tutup teks lama dengan kotak putih.
                        pg.add_redact_annot(rect, fill=(1, 1, 1))
                        pg.apply_redactions()
                    if dihapus or not teks_baru:
                        continue
                    if fs <= 0:
                        fs = min(max(4.0, rect.height * 0.85), 24)
                    fs = min(max(fs, 6.0), 36.0)
                    nama_font = 'hebo' if bold else ('heti' if italic else 'helv')
                    rgb = (0, 0, 0)
                    m = __import__('re').match(r'^#?([0-9a-fA-F]{6})$', str(warna).strip())
                    if m:
                        h = m.group(1)
                        rgb = (int(h[0:2], 16) / 255, int(h[2:4], 16) / 255, int(h[4:6], 16) / 255)
                    # Tulis teks baru di posisi sama.
                    pg.insert_text((rect.x0 + 1, rect.y1 - 2), teks_baru, fontsize=fs,
                                   fontname=nama_font, color=rgb)
                except Exception:
                    continue
            nama = f"edited_{uuid.uuid4().hex[:8]}.pdf"
            keluar = os.path.join(OUTPUT_FOLDER, nama)
            doc.save(keluar, garbage=3, deflate=True)
        finally:
            doc.close()
        return send_file(keluar, as_attachment=True, download_name=nama, mimetype='application/pdf'), 200
    except Exception as e:
        return jsonify({"error": f"Gagal menyimpan PDF: {e}"}), 500


@app.route('/api/word-to-pdf', methods=['POST'])
def word_to_pdf():
    import uuid
    if 'file' not in request.files:
        return jsonify({"error": "Tidak ada file yang diunggah."}), 400

    file = request.files['file']
    if not file.filename:
        return jsonify({"error": "Tidak ada file yang dipilih."}), 400

    filename = secure_filename(file.filename)
    if os.path.splitext(filename)[1].lower() not in ('.docx', '.doc'):
        return jsonify({"error": "File harus berformat Word (.doc/.docx)."}), 400

    tag = uuid.uuid4().hex[:8]
    base = filename.rsplit('.', 1)[0]
    input_path = os.path.join(UPLOAD_FOLDER, f"{tag}_{filename}")
    output_filename = base + '.pdf'
    output_path = os.path.join(OUTPUT_FOLDER, f"{tag}_{output_filename}")

    try:
        file.save(input_path)
        if os.path.getsize(input_path) > WORD2PDF_MAX_MB * 1024 * 1024:
            return jsonify({"error": f"Ukuran file maksimal {WORD2PDF_MAX_MB} MB."}), 400
        # Membutuhkan MS Word terinstal di OS server (Windows/Mac).
        # CoInitialize wajib karena Flask melayani request di thread terpisah.
        # Lock: Word COM tidak tahan dua konversi bersamaan.
        try:
            import pythoncom as _pc
            _pc.CoInitialize()
        except Exception:
            pass
        try:
            with _WORD2PDF_LOCK:
                convert(input_path, output_path)
        finally:
            try:
                import pythoncom as _pc2
                _pc2.CoUninitialize()
            except Exception:
                pass
        if not os.path.exists(output_path) or os.path.getsize(output_path) < 100:
            return jsonify({"error": "Konversi gagal menghasilkan PDF. Pastikan file Word valid."}), 500
        with open(output_path, 'rb') as fh:
            data = fh.read()
        resp = app.response_class(data, mimetype='application/pdf')
        resp.headers['Content-Disposition'] = f'attachment; filename="{output_filename}"'
        return resp
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({"error": f"Gagal mengonversi Word (pastikan MS Word terinstal di server): {str(e)}"}), 500
    finally:
        for _p in (input_path, output_path):
            try:
                if os.path.exists(_p):
                    os.remove(_p)
            except OSError:
                pass


# ===================== TAMBAHAN: FUNGSI RINGKASAN DIPAKAI ULANG =====================
# Ditenagai modul AI pelatihan_ringkas/ringkas_ai.py (splitter Indonesia +
# skor hibrida TF-IDF/PageRank yang bobotnya di-tuning lawan korpus dinas).
def generate_summary_from_text(text_content, rasio_ringkasan=0.35, min_kalimat=2, maks_kalimat=10):
    if not text_content or not text_content.strip():
        return "Teks tidak berisi konten untuk diringkas."

    import sys as _sys
    _root = os.path.dirname(os.path.abspath(__file__))
    if _root not in _sys.path:
        _sys.path.insert(0, _root)
    from pelatihan_ringkas.ringkas_ai import bersihkan_teks, split_kalimat_id
    from pelatihan_ringkas import ringkas_ai as _rai

    bersih = bersihkan_teks(text_content)
    total = len(split_kalimat_id(bersih))
    if total <= min_kalimat:
        return bersih.strip()

    jumlah = max(min_kalimat, min(maks_kalimat, round(total * rasio_ringkasan)))
    hasil, _stat = _rai.ringkas(text_content, cfg=_rai.muat_config(),
                                min_kalimat=min_kalimat, jumlah=jumlah)
    if not hasil.strip():
        return "Gagal membuat ringkasan. Teks mungkin terlalu pendek atau kurang bervariasi."
    return hasil
# ======================================================================================


# API 3: Ringkas Dokumen (dari file .docx, AI terlatih khusus Indonesia)
RINGKAS_MAX_MB = 20
RINGKAS_MODES = {'singkat', 'sedang', 'lengkap'}

@app.route('/api/summarize', methods=['POST'])
def summarize_document():
    import uuid
    if 'file' not in request.files:
        return jsonify({"error": "Tidak ada file yang diunggah"}), 400

    file = request.files['file']
    if not file.filename:
        return jsonify({"error": "Tidak ada file yang dipilih"}), 400

    filename = secure_filename(file.filename)
    if os.path.splitext(filename)[1].lower() != '.docx':
        return jsonify({"error": "File harus berformat Word (.docx)."}), 400

    mode = (request.form.get('mode') or 'sedang').strip().lower()
    if mode not in RINGKAS_MODES:
        mode = 'sedang'

    tag = uuid.uuid4().hex[:8]
    input_path = os.path.join(UPLOAD_FOLDER, f"{tag}_{filename}")
    try:
        file.save(input_path)
        if os.path.getsize(input_path) > RINGKAS_MAX_MB * 1024 * 1024:
            return jsonify({"error": f"Ukuran file maksimal {RINGKAS_MAX_MB} MB."}), 400

        # 1. Ekstrak teks lengkap: paragraf + tabel + header/footer
        import sys as _sys2
        _root2 = os.path.dirname(os.path.abspath(__file__))
        if _root2 not in _sys2.path:
            _sys2.path.insert(0, _root2)
        from pelatihan_ringkas import ringkas_ai as _rai
        try:
            text_content = _rai.ekstrak_docx(input_path)
        except Exception:
            return jsonify({"error": "File .docx rusak atau tidak valid."}), 400

        if not text_content.strip():
            return jsonify({"summary": "Dokumen tidak berisi teks untuk diringkas.",
                            "stat": {"total": 0, "diambil": 0, "kompresi": 0}})

        # 2. Ringkas dengan AI terlatih
        summary, stat = _rai.ringkas(text_content, mode=mode)
        stat['mode'] = mode
        return jsonify({"summary": summary, "stat": stat})
    except Exception as e:
        return jsonify({"error": f"Gagal memproses dokumen: {str(e)}"}), 500
    finally:
        try:
            if os.path.exists(input_path):
                os.remove(input_path)
        except OSError:
            pass


# API 3b: Jelaskan Dokumen (AI penjelas, bukan sekadar ringkasan)
# Mengembalikan: ringkasan + poin-poin penting berpandu skor + kata kunci
# topik + statistik dokumen, sehingga user paham ISI dokumen tersebut.
@app.route('/api/jelaskan', methods=['POST'])
def jelaskan_document():
    import uuid
    if 'file' not in request.files:
        return jsonify({"error": "Tidak ada file yang diunggah"}), 400

    file = request.files['file']
    if not file.filename:
        return jsonify({"error": "Tidak ada file yang dipilih"}), 400

    filename = secure_filename(file.filename)
    if os.path.splitext(filename)[1].lower() != '.docx':
        return jsonify({"error": "File harus berformat Word (.docx)."}), 400

    mode = (request.form.get('mode') or 'sedang').strip().lower()
    if mode not in RINGKAS_MODES:
        mode = 'sedang'

    tag = uuid.uuid4().hex[:8]
    input_path = os.path.join(UPLOAD_FOLDER, f"{tag}_{filename}")
    try:
        file.save(input_path)
        if os.path.getsize(input_path) > RINGKAS_MAX_MB * 1024 * 1024:
            return jsonify({"error": f"Ukuran file maksimal {RINGKAS_MAX_MB} MB."}), 400

        import sys as _sys3
        _root3 = os.path.dirname(os.path.abspath(__file__))
        if _root3 not in _sys3.path:
            _sys3.path.insert(0, _root3)
        from pelatihan_ringkas import ringkas_ai as _rai
        try:
            text_content = _rai.ekstrak_docx(input_path)
        except Exception:
            return jsonify({"error": "File .docx rusak atau tidak valid."}), 400

        if not text_content.strip():
            return jsonify({"error": "Dokumen tidak berisi teks untuk dijelaskan."}), 400

        bersih = _rai.bersihkan_teks(text_content)
        kalimat = _rai.split_kalimat_id(bersih)
        cfg = _rai.muat_config()

        # 1. Ringkasan inti
        summary, stat = _rai.ringkas(text_content, mode=mode)

        # 2. Poin-poin penting: 5 kalimat skor tertinggi + alasannya
        penjelasan_poin = []
        if kalimat:
            skor = _rai.skor_kalimat(kalimat, cfg)
            stemmer, stopwords = _rai._sastrawi()
            token_docs = [_rai._tokenize(k, stemmer, stopwords) for k in kalimat]
            kunci = cfg.get('kata_kunci', [])
            # Lewati artefak tabel (mis. "DISPOSISI | DISPOSISI") yang miskin kata unik
            kandidat = [i for i in range(len(kalimat))
                        if len(set(kalimat[i].lower().replace('|', ' ').split())) >= 4]
            if not kandidat:
                kandidat = list(range(len(kalimat)))
            urutan = sorted(kandidat, key=lambda i: skor[i], reverse=True)[:5]
            for i in urutan:
                alasan = []
                if re.search(r'\d', kalimat[i]):
                    alasan.append("memuat angka/data konkret")
                kk = [w for w in kunci if w in kalimat[i].lower()]
                if kk:
                    alasan.append("kata kunci dinas: " + ", ".join(kk[:3]))
                if i < max(len(kalimat) * 0.15, 1):
                    alasan.append("letak di awal dokumen (biasanya inti)")
                penjelasan_poin.append({
                    "kalimat": kalimat[i],
                    "skor": round(float(skor[i]), 3),
                    "mengapa_penting": "; ".join(alasan) if alasan else "keterkaitan tinggi dengan keseluruhan dokumen",
                })

        # 3. Kata kunci topik: kata berstem substantif tersering
        topik = []
        try:
            from collections import Counter
            stemmer, stopwords = _rai._sastrawi()
            semua = []
            for k in kalimat:
                semua.extend(_rai._tokenize(k, stemmer, stopwords))
            topik = [w for w, _ in Counter(semua).most_common(10)]
        except Exception:
            pass

        # 4. Penjelasan terstruktur: deskripsi + contoh/permisalan + penjelasan
        terstruktur = {"deskripsi": [], "contoh": [], "penjelasan_relevan": []}
        try:
            _n_amb = min(max(len(kalimat), 1), 9)
            _rank = sorted(range(len(kalimat)), key=lambda i: skor[i], reverse=True)[:_n_amb]
            _kal = [kalimat[i] for i in _rank]
            _sk = [float(skor[i]) for i in _rank]
            _st = _rai.struktur_penjelasan(_kal, _sk, cfg)
            terstruktur = {"deskripsi": _st.get("deskripsi", []),
                           "contoh": _st.get("contoh", []),
                           "penjelasan_relevan": _st.get("penjelasan", [])}
        except Exception:
            pass

        kata = len(bersih.split())
        return jsonify({
            "ringkasan": summary,
            "poin_penting": penjelasan_poin,
            "penjelasan_terstruktur": terstruktur,
            "kata_kunci_topik": topik,
            "stat": {
                "mode": mode,
                "total_kalimat": len(kalimat),
                "diambil": stat.get('diambil', 0),
                "total_kata": kata,
                "kompresi": stat.get('kompresi', 0),
            },
        })
    except Exception as e:
        return jsonify({"error": f"Gagal menjelaskan dokumen: {str(e)}"}), 500
    finally:
        try:
            if os.path.exists(input_path):
                os.remove(input_path)
        except OSError:
            pass


# ===================== TAMBAHAN: FITUR TUGAS SAYA (FIRESTORE - TAHAP 1) =====================
# Tabel MySQL: 'tugas'. ID baris berupa string (bukan angka lagi).

def _serialize_tugas(doc, sertakan_nama_pemberi=False, sertakan_nama_penerima=False):
    t = doc.to_dict()
    hasil = {
        'id': doc.id,
        'judul': t.get('judul'),
        'deskripsi': t.get('deskripsi'),
        'prioritas': t.get('prioritas', 'sedang'),
        'status': t.get('status', 'belum'),
        'user_id': t.get('user_id'),
        'assigned_by': t.get('assigned_by'),
        'deadline': t.get('deadline'),
        'deadline_tampil': None,
        'bisa_diedit': t.get('assigned_by') is None,
        'bisa_dihapus': t.get('assigned_by') is None,
        'pengajuan_status': t.get('pengajuan_status', 'belum_dikirim'),
        'file_hasil': t.get('file_hasil'),
        'file_template': t.get('file_template'),
        'catatan_user': t.get('catatan_user'),
        'catatan_atasan': t.get('catatan_atasan'),
        'diputus_pada': t.get('diputus_pada')
    }
    if t.get('deadline'):
        try:
            hasil['deadline_tampil'] = datetime.datetime.strptime(t.get('deadline'), '%Y-%m-%d').strftime('%d %b %Y')
        except Exception:
            hasil['deadline_tampil'] = t.get('deadline')

    if sertakan_nama_pemberi and t.get('assigned_by'):
        try:
            pemberi_doc = fdb.collection('users').document(t.get('assigned_by')).get()
            hasil['diberikan_oleh'] = pemberi_doc.to_dict().get('fullname') if pemberi_doc.exists else None
        except Exception:
            hasil['diberikan_oleh'] = None
    if sertakan_nama_penerima and t.get('user_id'):
        penerima_doc = fdb.collection('users').document(t.get('user_id')).get()
        hasil['diberikan_kepada'] = penerima_doc.to_dict().get('fullname') if penerima_doc.exists else 'Tidak diketahui'
        hasil['diberikan_kepada_id'] = t.get('user_id')
    return hasil


@app.route('/api/tugas', methods=['GET'])
@login_required
def get_tugas_list():
    user_id = session.get('user_id')
    status_filter = request.args.get('status', 'semua')

    query = fdb.collection('tugas').where('user_id', '==', user_id)
    docs = list(query.stream())

    hasil = [_serialize_tugas(d, sertakan_nama_pemberi=True) for d in docs]

    if status_filter and status_filter != 'semua':
        hasil = [h for h in hasil if h['status'] == status_filter]

    # Urutkan di Python: status lalu deadline (tanpa perlu index gabungan di MySQL)
    hasil.sort(key=lambda h: (h['status'], h['deadline'] or '9999-99-99'))
    return jsonify(hasil)


@app.route('/api/tugas', methods=['POST'])
@login_required
def add_tugas():
    data = request.get_json(silent=True) or request.form

    judul = (data.get('judul') or '').strip()
    deskripsi = (data.get('deskripsi') or '').strip()
    prioritas = data.get('prioritas', 'sedang')
    deadline_str = data.get('deadline')

    if not judul:
        return jsonify({"error": "Judul tugas wajib diisi."}), 400
    if prioritas not in ('rendah', 'sedang', 'tinggi'):
        prioritas = 'sedang'
    if deadline_str:
        try:
            datetime.datetime.strptime(deadline_str, '%Y-%m-%d')
        except ValueError:
            return jsonify({"error": "Format tanggal deadline tidak valid."}), 400

    try:
        new_ref = fdb.collection('tugas').document()
        new_ref.set({
            'user_id': session.get('user_id'),
            'assigned_by': None,
            'judul': judul,
            'deskripsi': deskripsi,
            'prioritas': prioritas,
            'status': 'belum',
            'deadline': deadline_str or None,
            'dibuat_pada': SERVER_TIMESTAMP
        })
        return jsonify({"message": "Tugas berhasil ditambahkan.", "id": new_ref.id}), 201
    except Exception as e:
        return jsonify({"error": f"Gagal menambahkan tugas: {str(e)}"}), 500


@app.route('/api/tugas/<string:tugas_id>', methods=['PUT'])
@login_required
def update_tugas(tugas_id):
    tugas_ref = fdb.collection('tugas').document(tugas_id)
    tugas_doc = tugas_ref.get()
    if not tugas_doc.exists:
        return jsonify({"error": "Tugas tidak ditemukan."}), 404

    tugas = tugas_doc.to_dict()
    if tugas.get('user_id') != session.get('user_id'):
        return jsonify({"error": "Anda tidak memiliki akses ke tugas ini."}), 403

    data = request.get_json(silent=True) or request.form
    hanya_boleh_ubah_status = tugas.get('assigned_by') is not None

    try:
        update_data = {}
        if not hanya_boleh_ubah_status:
            if 'judul' in data and data.get('judul', '').strip():
                update_data['judul'] = data.get('judul').strip()
            if 'deskripsi' in data:
                update_data['deskripsi'] = data.get('deskripsi', '').strip()
            if 'prioritas' in data and data.get('prioritas') in ('rendah', 'sedang', 'tinggi'):
                update_data['prioritas'] = data.get('prioritas')
            if 'deadline' in data:
                update_data['deadline'] = data.get('deadline') or None

        if 'status' in data and data.get('status') in ('belum', 'proses', 'selesai'):
            update_data['status'] = data.get('status')

        if update_data:
            tugas_ref.update(update_data)
        return jsonify({"message": "Tugas berhasil diperbarui."}), 200
    except Exception as e:
        return jsonify({"error": f"Gagal memperbarui tugas: {str(e)}"}), 500


@app.route('/api/tugas/<string:tugas_id>', methods=['DELETE'])
@login_required
def delete_tugas(tugas_id):
    tugas_ref = fdb.collection('tugas').document(tugas_id)
    tugas_doc = tugas_ref.get()
    if not tugas_doc.exists:
        return jsonify({"error": "Tugas tidak ditemukan."}), 404

    tugas = tugas_doc.to_dict()
    if tugas.get('user_id') != session.get('user_id'):
        return jsonify({"error": "Anda tidak memiliki akses ke tugas ini."}), 403
    if tugas.get('assigned_by') is not None:
        return jsonify({"error": "Tugas ini diberikan oleh atasan, hanya atasan yang bisa menghapusnya."}), 403

    try:
        tugas_ref.delete()
        return jsonify({"message": "Tugas berhasil dihapus."}), 200
    except Exception as e:
        return jsonify({"error": f"Gagal menghapus tugas: {str(e)}"}), 500


# ===================== KIRIM FILE HASIL TUGAS (staf -> persetujuan atasan) =====================
@app.route('/api/tugas/<string:tugas_id>/kirim', methods=['POST'])
@login_required
def kirim_hasil_tugas(tugas_id):
    tugas_ref = fdb.collection('tugas').document(tugas_id)
    tugas_doc = tugas_ref.get()
    if not tugas_doc.exists:
        return jsonify({"error": "Tugas tidak ditemukan."}), 404

    tugas = tugas_doc.to_dict()
    if tugas.get('user_id') != session.get('user_id'):
        return jsonify({"error": "Anda tidak memiliki akses ke tugas ini."}), 403

    if 'file' not in request.files or not request.files['file'].filename:
        return jsonify({"error": "Pilih file hasil tugas yang akan dikirim."}), 400

    file = request.files['file']
    if not _tugas_file_ok(file.filename):
        return jsonify({"error": "Jenis file tidak didukung. Gunakan PDF, Word, Excel, PowerPoint, gambar, TXT, CSV, ZIP, atau RAR."}), 400

    file.seek(0, os.SEEK_END)
    ukuran_mb = file.tell() / (1024 * 1024)
    file.seek(0)
    if ukuran_mb > MAX_TUGAS_FILE_MB:
        return jsonify({"error": f"Ukuran file maksimal {MAX_TUGAS_FILE_MB} MB."}), 400

    try:
        folder_tugas = os.path.join(TUGAS_UPLOAD_FOLDER, tugas_id)
        os.makedirs(folder_tugas, exist_ok=True)
        nama_asli = file.filename
        nama_simpan = datetime.datetime.now().strftime('%Y%m%d%H%M%S') + '_' + secure_filename(nama_asli)
        file.save(os.path.join(folder_tugas, nama_simpan))

        catatan_user = (request.form.get('catatan') or '').strip()
        tugas_ref.update({
            'pengajuan_status': 'menunggu',
            'file_hasil': {
                'nama_asli': nama_asli,
                'nama_simpan': nama_simpan,
                'dikirim_pada': datetime.datetime.now().strftime('%Y-%m-%d %H:%M')
            },
            'catatan_user': catatan_user or None,
            'catatan_atasan': None,
            'diputus_pada': None,
            'status': 'selesai'
        })
        return jsonify({"message": "File hasil berhasil dikirim ke atasan untuk persetujuan."}), 200
    except Exception as e:
        return jsonify({"error": f"Gagal mengirim file: {str(e)}"}), 500


@app.route('/api/tugas/<string:tugas_id>/unduh', methods=['GET'])
@login_required
def unduh_hasil_tugas(tugas_id):
    tugas_ref = fdb.collection('tugas').document(tugas_id)
    tugas_doc = tugas_ref.get()
    if not tugas_doc.exists:
        return jsonify({"error": "Tugas tidak ditemukan."}), 404

    tugas = tugas_doc.to_dict()
    saya = session.get('user_id')
    peran = session.get('user_role')
    boleh = (tugas.get('user_id') == saya
             or tugas.get('assigned_by') == saya
             or peran == 'admin')
    if not boleh:
        return jsonify({"error": "Anda tidak memiliki akses ke berkas ini."}), 403

    info = tugas.get('file_hasil') or {}
    if not info.get('nama_simpan'):
        return jsonify({"error": "Belum ada file hasil pada tugas ini."}), 404

    folder_tugas = os.path.join(TUGAS_UPLOAD_FOLDER, tugas_id)
    # ?inline=1 -> tampilkan di tab browser (PDF/gambar/teks); selain itu diunduh
    if request.args.get('inline') == '1':
        ekstensi = (info.get('nama_asli') or '').rsplit('.', 1)[-1].lower() if '.' in (info.get('nama_asli') or '') else ''
        if ekstensi in {'pdf', 'jpg', 'jpeg', 'png', 'txt', 'csv'}:
            return send_from_directory(folder_tugas, info['nama_simpan'], as_attachment=False)
    return send_from_directory(folder_tugas, info['nama_simpan'],
                               as_attachment=True,
                               download_name=info.get('nama_asli') or info['nama_simpan'])

@app.route('/api/tugas/<string:tugas_id>/template', methods=['GET'])
@login_required
def unduh_template_tugas(tugas_id):
    """Unduh/liat file template yang dilampirkan atasan saat memberi tugas."""
    tugas_ref = fdb.collection('tugas').document(tugas_id)
    tugas_doc = tugas_ref.get()
    if not tugas_doc.exists:
        return jsonify({"error": "Tugas tidak ditemukan."}), 404

    tugas = tugas_doc.to_dict()
    saya = session.get('user_id')
    peran = session.get('user_role')
    boleh = (tugas.get('user_id') == saya
             or tugas.get('assigned_by') == saya
             or peran == 'admin')
    if not boleh:
        return jsonify({"error": "Anda tidak memiliki akses ke berkas ini."}), 403

    info = tugas.get('file_template') or {}
    if not info.get('nama_simpan'):
        return jsonify({"error": "Tidak ada file template pada tugas ini."}), 404

    folder_tugas = os.path.join(TUGAS_UPLOAD_FOLDER, tugas_id)
    # ?inline=1 -> tampilkan di tab browser (PDF/gambar/teks); selain itu diunduh
    if request.args.get('inline') == '1':
        ekstensi = (info.get('nama_asli') or '').rsplit('.', 1)[-1].lower() if '.' in (info.get('nama_asli') or '') else ''
        if ekstensi in {'pdf', 'jpg', 'jpeg', 'png', 'txt', 'csv'}:
            return send_from_directory(folder_tugas, info['nama_simpan'], as_attachment=False)
    return send_from_directory(folder_tugas, info['nama_simpan'],
                               as_attachment=True,
                               download_name=info.get('nama_asli') or info['nama_simpan'])
# ===============================================================================================


# ===================== PERSETUJUAN ATASAN ATAS FILE HASIL STAF =====================
@app.route('/api/atasan/persetujuan', methods=['GET'])
@atasan_required
def daftar_persetujuan():
    status_filter = request.args.get('status', 'menunggu')
    try:
        if session.get('user_role') == 'admin':
            docs = list(fdb.collection('tugas').stream())
        else:
            docs = list(fdb.collection('tugas').where('assigned_by', '==', session.get('user_id')).stream())
        hasil = [_serialize_tugas(d, sertakan_nama_penerima=True) for d in docs]
        # Hanya tugas yang pernah dikirimi file hasil
        hasil = [h for h in hasil if h.get('pengajuan_status') != 'belum_dikirim']
        if status_filter and status_filter != 'semua':
            hasil = [h for h in hasil if h.get('pengajuan_status') == status_filter]
        # Menunggu dulu, lalu yang terbaru dikirim
        urutan = {'menunggu': 0, 'ditolak': 1, 'disetujui': 2}
        hasil.sort(key=lambda h: (urutan.get(h.get('pengajuan_status'), 9),
                                  (h.get('file_hasil') or {}).get('dikirim_pada', '') or ''),
                   reverse=False)
        return jsonify(hasil)
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route('/api/atasan/persetujuan/<string:tugas_id>', methods=['POST'])
@atasan_required
def putus_pengajuan(tugas_id):
    data = request.get_json(silent=True) or request.form
    keputusan = (data.get('keputusan') or '').strip().lower()
    catatan = (data.get('catatan') or '').strip()
    if keputusan not in ('setuju', 'tolak'):
        return jsonify({"error": "Keputusan harus 'setuju' atau 'tolak'."}), 400

    tugas_ref = fdb.collection('tugas').document(tugas_id)
    tugas_doc = tugas_ref.get()
    if not tugas_doc.exists:
        return jsonify({"error": "Tugas tidak ditemukan."}), 404

    tugas = tugas_doc.to_dict()
    saya = session.get('user_id')
    if session.get('user_role') != 'admin' and tugas.get('assigned_by') != saya:
        return jsonify({"error": "Anda tidak memiliki akses ke pengajuan ini."}), 403
    if tugas.get('pengajuan_status') != 'menunggu':
        return jsonify({"error": "Pengajuan ini sudah diproses sebelumnya."}), 400

    try:
        pembaruan = {
            'pengajuan_status': 'disetujui' if keputusan == 'setuju' else 'ditolak',
            'catatan_atasan': catatan or None,
            'diputus_pada': datetime.datetime.now().strftime('%Y-%m-%d %H:%M')
        }
        if keputusan == 'tolak':
            # Kembalikan ke proses agar staf bisa perbaiki & kirim ulang
            pembaruan['status'] = 'proses'
        tugas_ref.update(pembaruan)

        if keputusan == 'tolak':
            # Kirim notifikasi penolakan ke Telegram staf (tidak menggagalkan respons jika gagal)
            try:
                staf_doc = fdb.collection('users').document(tugas.get('user_id')).get()
                staf = staf_doc.to_dict() if staf_doc.exists else None
                atasan_doc = fdb.collection('users').document(saya).get()
                nama_atasan = atasan_doc.to_dict().get('fullname', 'Atasan') if atasan_doc.exists else 'Atasan'
                if staf and staf.get('telegram_chat_id'):
                    public_url = get_cloudflared_url()
                    task_link = f'{public_url}/tugas-saya.html' if public_url else 'Aplikasi Portal (buka di browser)'
                    file_info = tugas.get('file_hasil') or {}
                    pesan_tolak = (
                        f'❌ <b>HASIL TUGAS DITOLAK</b>\n\n'
                        f'📝 <b>Judul:</b> {tugas.get("judul", "-")}\n'
                        f'📎 <b>Berkas:</b> {file_info.get("nama_asli", "-")}\n'
                        f'💬 <b>Alasan penolakan:</b> {catatan or "-"}\n'
                        f'👤 <b>Ditolak oleh:</b> {nama_atasan}\n'
                        f'⏰ <b>Waktu:</b> {datetime.datetime.now().strftime("%d %B %Y %H:%M")}\n\n'
                        f'Silakan perbaiki lalu kirim ulang.\n'
                        f'<a href="{task_link}">Buka aplikasi untuk mengirim ulang</a>'
                    )
                    send_telegram_message(staf.get('telegram_chat_id'), pesan_tolak)
                else:
                    print('[TELEGRAM] Staf belum menghubungkan akun Telegram, notifikasi penolakan dilewati')
            except Exception as e:
                print(f'[TELEGRAM] Gagal kirim notifikasi penolakan: {e}')

        pesan = 'Hasil tugas disetujui.' if keputusan == 'setuju' else 'Hasil tugas ditolak dan dikembalikan ke staf.'
        return jsonify({"message": pesan}), 200
    except Exception as e:
        return jsonify({"error": f"Gagal memproses persetujuan: {str(e)}"}), 500
# ===================================================================================


@app.route('/USER/tugas-saya.html')
@login_required
def serve_user_tugas():
    return render_template('USER/tugas-saya.html', user_name=session.get('user_name', 'Pengguna'), user_id=session.get('user_id'))

# ===============================================================================================


# ===================== TAMBAHAN: ATASAN MEMBERI TUGAS KE STAF (FIRESTORE) =====================

@app.route('/api/atasan/staf', methods=['GET'])
@atasan_required
def get_staf_tim():
    atasan_id = session.get('user_id')
    try:
        if session.get('user_role') == 'admin':
            docs = fdb.collection('users').where('role', '==', 'user').stream()
        else:
            docs = fdb.collection('users').where('superior_id', '==', atasan_id).stream()
        return jsonify([{'id': d.id, 'fullname': d.to_dict().get('fullname'), 'email': d.to_dict().get('email')} for d in docs])
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route('/api/atasan/tugas', methods=['GET'])
@atasan_required
def get_tugas_yang_diberikan():
    atasan_id = session.get('user_id')
    try:
        docs = list(fdb.collection('tugas').where('assigned_by', '==', atasan_id).stream())
        hasil = [_serialize_tugas(d, sertakan_nama_penerima=True) for d in docs]
        hasil.sort(key=lambda h: h['status'])
        return jsonify(hasil)
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route('/api/atasan/tugas', methods=['POST'])
@atasan_required
def beri_tugas():
    data = request.get_json(silent=True) or request.form

    target_user_id = data.get('user_id')
    judul = (data.get('judul') or '').strip()
    deskripsi = (data.get('deskripsi') or '').strip()
    prioritas = data.get('prioritas', 'sedang')
    deadline_str = data.get('deadline')

    if not target_user_id or not judul:
        return jsonify({"error": "Staf tujuan dan judul tugas wajib diisi."}), 400

    target_doc = fdb.collection('users').document(target_user_id).get()
    if not target_doc.exists:
        return jsonify({"error": "Staf tujuan tidak ditemukan."}), 404
    target_user = target_doc.to_dict()

    if session.get('user_role') == 'atasan' and target_user.get('superior_id') != session.get('user_id'):
        return jsonify({"error": "Anda hanya bisa memberi tugas ke staf di tim Anda sendiri."}), 403

    if prioritas not in ('rendah', 'sedang', 'tinggi'):
        prioritas = 'sedang'
    if deadline_str:
        try:
            datetime.datetime.strptime(deadline_str, '%Y-%m-%d')
        except ValueError:
            return jsonify({"error": "Format tanggal deadline tidak valid."}), 400

    try:
        new_task_ref = fdb.collection('tugas').document()
        tugas_id_baru = new_task_ref.id

        # Lampiran template dari atasan (opsional): tersimpan per-tugas agar
        # bawahan bisa melihat & mengunduhnya dari halaman Tugas Saya.
        file_template = None
        tpl = request.files.get('template')
        if tpl and tpl.filename:
            if not _tugas_file_ok(tpl.filename):
                return jsonify({"error": "Jenis file template tidak didukung. Gunakan PDF, Word, Excel, PowerPoint, gambar, TXT, CSV, ZIP, atau RAR."}), 400
            tpl.seek(0, os.SEEK_END)
            if tpl.tell() / (1024 * 1024) > MAX_TUGAS_FILE_MB:
                return jsonify({"error": f"Ukuran file template maksimal {MAX_TUGAS_FILE_MB} MB."}), 400
            tpl.seek(0)
            folder_baru = os.path.join(TUGAS_UPLOAD_FOLDER, tugas_id_baru)
            os.makedirs(folder_baru, exist_ok=True)
            nama_simpan_tpl = 'template_' + datetime.datetime.now().strftime('%Y%m%d%H%M%S') + '_' + secure_filename(tpl.filename)
            tpl.save(os.path.join(folder_baru, nama_simpan_tpl))
            file_template = {'nama_asli': tpl.filename, 'nama_simpan': nama_simpan_tpl}

        new_task_ref.set({
            'user_id': target_user_id,
            'assigned_by': session.get('user_id'),
            'judul': judul,
            'deskripsi': deskripsi,
            'prioritas': prioritas,
            'status': 'belum',
            'deadline': deadline_str or None,
            'file_template': file_template,
            'dibuat_pada': SERVER_TIMESTAMP
        })
        
        # Kirim notifikasi Telegram ke staf yang mendapat tugas
        assigned_by_doc = fdb.collection('users').document(session.get('user_id')).get()
        assigned_by_name = assigned_by_doc.to_dict().get('fullname', 'Atasan') if assigned_by_doc.exists else 'Atasan'
        
        task_data = {
            'judul': judul,
            'deskripsi': deskripsi,
            'prioritas': prioritas,
            'deadline': deadline_str,
            'template': (file_template or {}).get('nama_asli')
        }
        
        # Debug: check if target_user has telegram_chat_id
        chat_id = target_user.get('telegram_chat_id')
        print(f'[TELEGRAM DEBUG] Target user: {target_user.get("fullname")}, chat_id: {chat_id}')
        
        if chat_id:
            send_task_notification_telegram(target_user, task_data, assigned_by_name)
        else:
            print('[TELEGRAM] Target user belum menghubungkan akun Telegram')
        
        return jsonify({"message": f"Tugas berhasil diberikan ke {target_user.get('fullname')}."}), 201
    except Exception as e:
        return jsonify({"error": f"Gagal memberi tugas: {str(e)}"}), 500


@app.route('/api/atasan/tugas/<string:tugas_id>', methods=['DELETE'])
@atasan_required
def batalkan_tugas(tugas_id):
    tugas_ref = fdb.collection('tugas').document(tugas_id)
    tugas_doc = tugas_ref.get()
    if not tugas_doc.exists:
        return jsonify({"error": "Tugas tidak ditemukan."}), 404

    tugas = tugas_doc.to_dict()
    if session.get('user_role') != 'admin' and tugas.get('assigned_by') != session.get('user_id'):
        return jsonify({"error": "Anda tidak memiliki akses ke tugas ini."}), 403

    try:
        tugas_ref.delete()
        # Bersihkan berkas template & hasil agar tidak jadi sampah di server
        try:
            import shutil
            folder_tugas = os.path.join(TUGAS_UPLOAD_FOLDER, tugas_id)
            if os.path.isdir(folder_tugas):
                shutil.rmtree(folder_tugas, ignore_errors=True)
        except Exception:
            pass
        return jsonify({"message": "Tugas berhasil dibatalkan."}), 200
    except Exception as e:
        return jsonify({"error": f"Gagal membatalkan tugas: {str(e)}"}), 500


@app.route('/atasan/beri-tugas.html')
@login_required
def serve_atasan_beri_tugas():
    if session.get('user_role') != 'atasan':
        return "Akses Ditolak", 403
    return render_template('atasan/beri-tugas.html', user_name=session.get('user_name', 'Atasan'), user_id=session.get('user_id'))


@app.route('/atasan/tugas-saya.html')
@login_required
def serve_atasan_tugas_saya():
    if session.get('user_role') != 'atasan':
        return "Akses Ditolak", 403
    # Atasan paling tinggi (tanpa atasan langsung) tidak memiliki Tugas Saya:
    # arahkan kembali ke dashboard, bukan halaman error.
    saya = fdb.collection('users').document(session.get('user_id')).get()
    if saya.exists and not (saya.to_dict() or {}).get('superior_id'):
        return redirect(url_for('serve_atasan_dashboard'))
    return render_template('atasan/tugas-saya.html', user_name=session.get('user_name', 'Atasan'), user_id=session.get('user_id'))


@app.route('/atasan/profile.html')
@login_required
def serve_atasan_profile():
    if session.get('user_role') != 'atasan':
        return "Akses Ditolak", 403
    return render_template('atasan/profile.html', user_name=session.get('user_name', 'Atasan'), user_email=session.get('user_email', ''))


# ===================== KONFIGURASI FIREBASE (SUDAH TIDAK DIPAKAI) =====================
# Frontend tidak lagi memakai Firebase Web SDK. Halaman tugas sekarang mengambil
# data lewat REST API (/api/tugas dan /api/atasan/tugas) dengan polling 10 detik.
# Endpoint /api/firebase-config dihapus agar config Firebase tidak lagi terekspos.
# =============================================================================================





# ==========================================
# API UNTUK OTENTIKASI PENGGUNA
# ==========================================

@app.route('/login', methods=['POST'])
def handle_login():
    email = request.form.get('email')
    password = request.form.get('password')

    # ===== FIRESTORE (Tahap 1) =====
    users_query = fdb.collection('users').where('email', '==', email).limit(1).stream()
    user_doc = next(users_query, None)

    if user_doc and check_password_hash(user_doc.to_dict().get('password_hash', ''), password):
        user = user_doc.to_dict()
        session['user_id'] = user_doc.id
        session['user_name'] = user.get('fullname')
        session['user_email'] = user.get('email')
        session['user_role'] = user.get('role')

        role = user.get('role')
        if role == 'admin':
            redirect_url = url_for('serve_admin_dashboard')
        elif role == 'atasan':
            redirect_url = url_for('serve_atasan_dashboard')
        else:  # 'user'
            redirect_url = url_for('serve_user_dashboard')

        return jsonify({"message": "Login berhasil!", "redirect_url": redirect_url})

    return jsonify({"error": "Email atau kata sandi salah"}), 401

@app.route('/register', methods=['POST'])
def handle_register():
    fullname = request.form.get('fullname')
    email = request.form.get('email')
    password = request.form.get('password')

    if not all([fullname, email, password]):
        return jsonify({"error": "Semua kolom harus diisi."}), 400

    # ===== FIRESTORE (Tahap 1) =====
    existing = list(fdb.collection('users').where('email', '==', email).limit(1).stream())
    if existing:
        return jsonify({"error": "Email sudah terdaftar."}), 409

    hashed_password = generate_password_hash(password)
    fdb.collection('users').document().set({
        'fullname': fullname,
        'email': email,
        'password_hash': hashed_password,
        'role': 'user',
        'superior_id': None
    })

    return jsonify({"message": "Registrasi berhasil! Anda akan diarahkan ke halaman login.", "redirect_url": url_for('serve_login_page')})

@app.route('/api/users', methods=['GET'])
@admin_required
def get_users():
    try:
        docs = fdb.collection('users').stream()
        users_list = []
        for doc in docs:
            u = doc.to_dict()
            pantau, kirim = _hak_disp_efektif(u, u.get('role'), doc.id)
            users_list.append({
                'id': doc.id,
                'fullname': u.get('fullname'),
                'email': u.get('email'),
                'role': u.get('role'),
                'jabatan_atasan': u.get('jabatan_atasan') or '',
                'bagian': u.get('bagian') or '',
                'superior_id': u.get('superior_id'),
                'disp_pantau': pantau,
                'disp_kirim': kirim
            })
        return jsonify(users_list)
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route('/api/direktori', methods=['GET'])
@login_required
def get_direktori():
    """Daftar staf untuk halaman Direktori (data dikelola admin via Kelola Pengguna).

    Hanya field publik: nama, email, peran, dan nama atasan langsung.
    """
    try:
        docs = list(fdb.collection('users').stream())
        cache_nama = {d.id: (d.to_dict().get('fullname') or '-') for d in docs}
        hasil = []
        for d in docs:
            u = d.to_dict()
            hasil.append({
                'id': d.id,
                'fullname': u.get('fullname'),
                'email': u.get('email'),
                'role': u.get('role'),
                'jabatan_atasan': u.get('jabatan_atasan') or '',
                'bagian': u.get('bagian') or '',
                'atasan': cache_nama.get(u.get('superior_id')) if u.get('superior_id') else '-'
            })
        hasil.sort(key=lambda h: (h.get('fullname') or '').lower())
        return jsonify(hasil)
    except Exception as e:
        return jsonify({"error": str(e)}), 500


# ===================== STRUKTUR ORGANISASI (kotak bagan direktori) =====================
# Hanya admin yang boleh tambah/hapus kotak. Semua peran yang login boleh membaca.
WARNA_STRUKTUR = {'merah', 'kuning', 'biru', 'teal', 'nila', 'ungu'}

def _seed_struktur_awal():
    """Isi awal bagan sesuai struktur yang selama ini tampil statis. Sekali saja."""
    data = [
        {'nama': 'Jenderal Budi', 'jabatan': 'Komandan', 'jabatan_panjang': 'Komandan Utama',
         'email': 'jend.budi@perusahaan.co.id', 'telepon': '+62 811-0000-0001', 'warna': 'merah', 'parent_nama': None},
        {'nama': 'Jenderal Ahmad', 'jabatan': 'Wakil Komandan', 'jabatan_panjang': 'Wakil Komandan',
         'email': 'jend.ahmad@perusahaan.co.id', 'telepon': '+62 811-0000-0002', 'warna': 'kuning', 'parent_nama': 'Jenderal Budi'},
        {'nama': 'Kolonel Siti', 'jabatan': 'Asintelops', 'jabatan_panjang': 'Asisten Intelijen & Operasi',
         'email': 'kol.siti@perusahaan.co.id', 'telepon': '+62 812-0000-0003', 'warna': 'biru', 'parent_nama': 'Jenderal Ahmad'},
        {'nama': 'Kolonel Dewi', 'jabatan': 'Asiminlog', 'jabatan_panjang': 'Asisten Administrasi & Logistik',
         'email': 'kol.dewi@perusahaan.co.id', 'telepon': '+62 812-0000-0004', 'warna': 'nila', 'parent_nama': 'Jenderal Ahmad'},
        {'nama': 'Mayor Ridwan', 'jabatan': 'Dansat Certan', 'jabatan_panjang': 'Dansat Cyber Pertahanan',
         'email': 'myr.ridwan@perusahaan.co.id', 'telepon': '+62 813-0000-0005', 'warna': 'teal', 'parent_nama': 'Kolonel Siti'},
        {'nama': 'Mayor Prasetyo', 'jabatan': 'Dansat Linta', 'jabatan_panjang': 'Dansat Lintas Udara',
         'email': 'myr.prasetyo@perusahaan.co.id', 'telepon': '+62 813-0000-0006', 'warna': 'teal', 'parent_nama': 'Kolonel Siti'},
        {'nama': 'Mayor Nurhaliza', 'jabatan': 'Dansat Indigi', 'jabatan_panjang': 'Dansat Industri Digital',
         'email': 'myr.nurhaliza@perusahaan.co.id', 'telepon': '+62 813-0000-0007', 'warna': 'ungu', 'parent_nama': 'Kolonel Dewi'},
        {'nama': 'Mayor Lestari', 'jabatan': 'Dansat Deltan', 'jabatan_panjang': 'Dansat Delta Tani',
         'email': 'myr.lestari@perusahaan.co.id', 'telepon': '+62 813-0000-0008', 'warna': 'ungu', 'parent_nama': 'Kolonel Dewi'},
    ]
    id_per_nama = {}
    for i, d in enumerate(data):
        ref = fdb.collection('struktur').document()
        ref.set({
            'nama': d['nama'],
            'jabatan': d['jabatan'],
            'jabatan_panjang': d.get('jabatan_panjang') or d['jabatan'],
            'email': d.get('email'),
            'telepon': d.get('telepon'),
            'warna': d.get('warna') if d.get('warna') in WARNA_STRUKTUR else 'biru',
            'parent_id': id_per_nama.get(d['parent_nama']),
            'urutan': i
        })
        id_per_nama[d['nama']] = ref.id


@app.route('/api/struktur', methods=['GET'])
@login_required
def get_struktur():
    try:
        docs = list(fdb.collection('struktur').stream())
        if not docs:
            _seed_struktur_awal()
            docs = list(fdb.collection('struktur').stream())
        hasil = []
        for d in docs:
            s = d.to_dict()
            hasil.append({
                'id': d.id,
                'nama': s.get('nama'),
                'jabatan': s.get('jabatan'),
                'jabatan_panjang': s.get('jabatan_panjang') or s.get('jabatan'),
                'email': s.get('email'),
                'telepon': s.get('telepon'),
                'warna': s.get('warna', 'biru'),
                'parent_id': s.get('parent_id'),
                'urutan': s.get('urutan', 0),
                'user_id': s.get('user_id')
            })
        hasil.sort(key=lambda h: (h['nama'] or '').lower())
        return jsonify(hasil)
    except Exception as e:
        return jsonify({"error": str(e)}), 500


def _struktur_cek_siklus(node_id, calon_parent_id):
    """Cegah hierarki melingkar: calon atasan tak boleh diri sendiri/descendant."""
    if not calon_parent_id:
        return True
    if calon_parent_id == node_id:
        return False
    seen, cur = set(), calon_parent_id
    while cur:
        if cur == node_id or cur in seen:
            return False
        seen.add(cur)
        doc = fdb.collection('struktur').document(cur).get()
        if not doc.exists:
            break
        cur = doc.to_dict().get('parent_id')
    return True


@app.route('/api/struktur', methods=['POST'])
@admin_required
def tambah_struktur():
    data = request.get_json(silent=True) or request.form
    nama = (data.get('nama') or '').strip()
    jabatan = (data.get('jabatan') or '').strip()
    user_id = (data.get('user_id') or '').strip() or None
    # Tautkan ke akun terdaftar: auto-isi dari data pengguna bila kolom kosong
    akun = None
    if user_id:
        akun = fdb.collection('users').document(user_id).get()
        if not akun.exists:
            return jsonify({"error": "Pengguna yang dipilih tidak ditemukan."}), 404
        akun = akun.to_dict()
        if not nama:
            nama = akun.get('fullname') or ''
        if not jabatan:
            jabatan = akun.get('jabatan_atasan') or akun.get('bagian') or (
                'Atasan' if akun.get('role') == 'atasan' else 'Staf')
        for _dup in fdb.collection('struktur').where('user_id', '==', user_id).stream():
            siapa = akun.get('fullname') or nama
            return jsonify({"error": f"{siapa} sudah punya kotak di bagan. Ubah kotaknya, bukan tambah baru."}), 409
    if not nama or not jabatan:
        return jsonify({"error": "Nama dan jabatan wajib diisi."}), 400
    warna = (data.get('warna') or 'biru').strip().lower()
    if warna not in WARNA_STRUKTUR:
        warna = 'biru'
    parent_id = (data.get('parent_id') or '').strip() or None
    if parent_id and not fdb.collection('struktur').document(parent_id).get().exists:
        return jsonify({"error": "Atasan yang dipilih tidak ditemukan."}), 404
    try:
        ref = fdb.collection('struktur').document()
        ref.set({
            'nama': nama,
            'jabatan': jabatan,
            'jabatan_panjang': (data.get('jabatan_panjang') or '').strip() or jabatan,
            'email': (data.get('email') or '').strip() or (akun.get('email') if akun else None),
            'telepon': (data.get('telepon') or '').strip() or None,
            'warna': warna,
            'parent_id': parent_id,
            'user_id': user_id
        })
        return jsonify({"message": f"Kotak '{nama}' berhasil ditambahkan ke bagan."}), 201
    except Exception as e:
        return jsonify({"error": f"Gagal menambah kotak: {str(e)}"}), 500


@app.route('/api/struktur/<string:node_id>', methods=['PUT'])
@admin_required
def ubah_struktur(node_id):
    """Ubah kotak: nama/jabatan/kontak/warna/urutan + PINDAH atasan langsung."""
    ref = fdb.collection('struktur').document(node_id)
    doc = ref.get()
    if not doc.exists:
        return jsonify({"error": "Kotak tidak ditemukan."}), 404
    data = request.get_json(silent=True) or request.form
    lama = doc.to_dict()
    nama = (data.get('nama') or lama.get('nama') or '').strip()
    jabatan = (data.get('jabatan') or lama.get('jabatan') or '').strip()
    if not nama or not jabatan:
        return jsonify({"error": "Nama dan jabatan wajib diisi."}), 400
    warna = (data.get('warna') or lama.get('warna') or 'biru').strip().lower()
    if warna not in WARNA_STRUKTUR:
        warna = 'biru'
    parent_id = (data.get('parent_id') or '').strip() or None
    if parent_id and not fdb.collection('struktur').document(parent_id).get().exists:
        return jsonify({"error": "Atasan yang dipilih tidak ditemukan."}), 404
    if parent_id != lama.get('parent_id') and not _struktur_cek_siklus(node_id, parent_id):
        return jsonify({"error": "Atasan tidak valid: akan membuat hierarki melingkar."}), 400
    try:
        ref.update({
            'nama': nama,
            'jabatan': jabatan,
            'jabatan_panjang': (data.get('jabatan_panjang') or lama.get('jabatan_panjang') or jabatan).strip(),
            'email': (data.get('email') or lama.get('email') or '').strip() or None,
            'telepon': (data.get('telepon') or lama.get('telepon') or '').strip() or None,
            'warna': warna,
            'parent_id': parent_id
        })
        return jsonify({"message": f"Kotak '{nama}' berhasil diperbarui."}), 200
    except Exception as e:
        return jsonify({"error": f"Gagal memperbarui kotak: {str(e)}"}), 500


@app.route('/api/struktur/sinkron-atasan', methods=['POST'])
@admin_required
def sinkron_struktur_dari_atasan():
    """Bangun/selaraskan kotak bagan dari data akun atasan terdaftar.

    Hierarki mengikuti superior_id masing-masing atasan, sehingga bagan
    selalu sesuai struktur atasan langsung. Kotak manual (tanpa user_id)
    tidak dihapus/diubah.
    """
    try:
        atasan_docs = list(fdb.collection('users').where('role', '==', 'atasan').stream())
        if not atasan_docs:
            return jsonify({"error": "Belum ada akun atasan terdaftar."}), 400
        node_per_user = {}
        for n in fdb.collection('struktur').stream():
            uid = (n.to_dict() or {}).get('user_id')
            if uid:
                node_per_user[uid] = n.id
        dibuat, diperbarui, dilewati = 0, 0, []
        # Pass berulang agar induk selalu dibuat sebelum anaknya
        sisa = {d.id: (d.to_dict() or {}) for d in atasan_docs}
        for _ in range(len(sisa) + 1):
            maju = False
            for uid, u in list(sisa.items()):
                sup = u.get('superior_id')
                if sup and sup in sisa:
                    continue  # tunggu induknya dulu
                parent_node = node_per_user.get(sup) if sup else None
                nama = u.get('fullname') or '-'
                jabatan = u.get('jabatan_atasan') or u.get('bagian') or 'Atasan'
                if uid in node_per_user:
                    fdb.collection('struktur').document(node_per_user[uid]).update({
                        'nama': nama, 'jabatan': jabatan,
                        'email': u.get('email') or None, 'parent_id': parent_node})
                    diperbarui += 1
                else:
                    ref = fdb.collection('struktur').document()
                    ref.set({'nama': nama, 'jabatan': jabatan,
                             'jabatan_panjang': jabatan, 'email': u.get('email') or None,
                             'telepon': None, 'warna': 'biru', 'parent_id': parent_node,
                             'user_id': uid})
                    node_per_user[uid] = ref.id
                    dibuat += 1
                del sisa[uid]
                maju = True
            if not sisa or not maju:
                break
        for uid, u in sisa.items():
            dilewati.append(u.get('fullname') or uid)
        pesan = f"Sinkron selesai: {dibuat} kotak dibuat, {diperbarui} diperbarui."
        if dilewati:
            pesan += f" Dilewati (atasan di atasnya bukan atasan terdaftar): {', '.join(dilewati)}."
        return jsonify({"message": pesan, "dibuat": dibuat, "diperbarui": diperbarui,
                        "dilewati": dilewati}), 200
    except Exception as e:
        return jsonify({"error": f"Gagal sinkronisasi: {str(e)}"}), 500


@app.route('/api/struktur/<string:node_id>', methods=['DELETE'])
@admin_required
def hapus_struktur(node_id):
    ref = fdb.collection('struktur').document(node_id)
    doc = ref.get()
    if not doc.exists:
        return jsonify({"error": "Kotak tidak ditemukan."}), 404
    try:
        # Anak-anaknya naik menggantikan posisi kotak yang dihapus
        induk_baru = doc.to_dict().get('parent_id')
        for anak in fdb.collection('struktur').where('parent_id', '==', node_id).stream():
            anak.reference.update({'parent_id': induk_baru})
        ref.delete()
        return jsonify({"message": "Kotak berhasil dihapus dari bagan."}), 200
    except Exception as e:
        return jsonify({"error": f"Gagal menghapus kotak: {str(e)}"}), 500
# ===================================================================================


@app.route('/api/superiors', methods=['GET'])
@admin_required
def get_superiors():
    try:
        docs = fdb.collection('users').where('role', '==', 'atasan').stream()
        hasil=[]
        for d in docs:
            u=d.to_dict()
            hasil.append({'id': d.id, 'fullname': u.get('fullname'), 'jabatan_atasan': u.get('jabatan_atasan') or ''})
        return jsonify(hasil)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/api/user/profile', methods=['GET'])
@login_required
def get_user_profile():
    user_id = session.get('user_id')
    user_ref = fdb.collection('users').document(user_id)
    user_doc = user_ref.get()
    if not user_doc.exists:
        return jsonify({"error": "Pengguna tidak ditemukan."}), 404
    user = user_doc.to_dict()
    pantau, kirim = _hak_disp_efektif(user, user.get('role', ''), user_id)
    _jabatan_tampil = (user.get('jabatan_atasan') or user.get('bagian')
                       or {'admin': 'Administrator', 'atasan': 'Atasan'}.get(user.get('role', ''), 'Staf'))
    _atasan_nama = ''
    try:
        _sup = user.get('superior_id')
        if _sup:
            _sd = fdb.collection('users').document(_sup).get()
            if _sd.exists:
                _atasan_nama = (_sd.to_dict() or {}).get('fullname') or ''
    except Exception:
        pass
    return jsonify({
        'nama': user.get('fullname', ''),
        'role': user.get('role', ''),
        'bagian': user.get('bagian') or '',
        'jabatan_tampil': _jabatan_tampil,
        'atasan_nama': _atasan_nama,
        'tingkat': tingkat_of(user_id),
        'disp_pantau': pantau,
        'disp_kirim': kirim,
        'punya_pin_disposisi': bool(user.get('pin_disposisi_hash')),
        'nrp': user.get('nrp', ''),
        'asal': user.get('asal', ''),
        'gender': user.get('gender', ''),
        'email': user.get('email', ''),
        'telepon': user.get('telepon', ''),
        'foto': user.get('foto', ''),
        'telegram_chat_id': user.get('telegram_chat_id', ''),
        'punya_atasan_langsung': bool(user.get('superior_id'))
    }), 200


@app.route('/api/user/profile', methods=['POST'])
@login_required
def update_user_profile():
    user_id = session.get('user_id')
    user_ref = fdb.collection('users').document(user_id)
    user_doc = user_ref.get()
    if not user_doc.exists:
        return jsonify({"error": "Pengguna tidak ditemukan."}), 404

    try:
        # Handle form data (including file upload)
        if request.content_type and request.content_type.startswith('multipart/form-data'):
            nama = request.form.get('nama', '').strip()
            nrp = request.form.get('nrp', '').strip()
            asal = request.form.get('asal', '').strip()
            gender = request.form.get('gender', '').strip()
            email = request.form.get('email', '').strip()
            telepon = request.form.get('telepon', '').strip()
            tg_chat = request.form.get('telegram_chat_id', '').strip()
            
            update_data = {}
            if nama: update_data['fullname'] = nama
            if nrp: update_data['nrp'] = nrp
            if asal: update_data['asal'] = asal
            if gender: update_data['gender'] = gender
            if email: update_data['email'] = email
            if telepon: update_data['telepon'] = telepon
            if tg_chat: update_data['telegram_chat_id'] = tg_chat
            
            # Handle photo upload
            if 'foto' in request.files:
                file = request.files['foto']
                if file and file.filename:
                    # Convert to base64 for storage in MySQL
                    import base64
                    file_data = file.read()
                    file_b64 = base64.b64encode(file_data).decode('utf-8')
                    mime_type = file.content_type or 'image/jpeg'
                    update_data['foto'] = f'data:{mime_type};base64,{file_b64}'
            
            if update_data:
                user_ref.update(update_data)
                # Update session name if changed
                if nama:
                    session['user_name'] = nama
                return jsonify({"message": "Profil berhasil diperbarui."}), 200
            else:
                return jsonify({"error": "Tidak ada data yang diubah."}), 400
        else:
            return jsonify({"error": "Content-Type harus multipart/form-data"}), 400
            
    except Exception as e:
        return jsonify({"error": f"Gagal memperbarui profil: {str(e)}"}), 500


@app.route('/api/change-password', methods=['POST'])
@login_required
def change_password():
    user_id = session.get('user_id')
    old_password = request.form.get('old_password')
    new_password = request.form.get('new_password')
    confirm_new_password = request.form.get('confirm_new_password')

    if not all([old_password, new_password, confirm_new_password]):
        return jsonify({"error": "Semua kolom harus diisi."}), 400

    if new_password != confirm_new_password:
        return jsonify({"error": "Kata sandi baru dan konfirmasi tidak cocok."}), 400

    if len(new_password) < 4:
        return jsonify({"error": "Kata sandi baru minimal 4 karakter."}), 400

    user_ref = fdb.collection('users').document(user_id)
    user_doc = user_ref.get()
    if not user_doc.exists:
        session.clear()
        return jsonify({"error": "Pengguna tidak ditemukan."}), 401

    user = user_doc.to_dict()
    if not check_password_hash(user.get('password_hash', ''), old_password):
        return jsonify({"error": "Kata sandi lama salah."}), 401

    user_ref.update({'password_hash': generate_password_hash(new_password)})

    return jsonify({"message": "Kata sandi berhasil diubah."}), 200


@app.route('/api/user/pin-disposisi', methods=['POST'])
@login_required
def atur_pin_disposisi():
    """Atur/ubah PIN disposisi lapis-dua (dipakai staf persuratan untuk
    verifikasi aksi Ubah & Hapus surat). Disimpan sebagai hash, min 4 karakter."""
    user_id = session.get('user_id')
    pin_lama = (request.form.get('pin_lama') or '').strip()
    pin_baru = (request.form.get('pin_baru') or '').strip()
    konfirmasi = (request.form.get('konfirmasi_pin') or '').strip()
    if not pin_baru or not konfirmasi:
        return jsonify({"error": "PIN baru dan konfirmasi harus diisi."}), 400
    if pin_baru != konfirmasi:
        return jsonify({"error": "PIN baru dan konfirmasi tidak cocok."}), 400
    if len(pin_baru) < 4:
        return jsonify({"error": "PIN minimal 4 karakter."}), 400
    user_ref = fdb.collection('users').document(user_id)
    user_doc = user_ref.get()
    if not user_doc.exists:
        return jsonify({"error": "Pengguna tidak ditemukan."}), 404
    lama_hash = (user_doc.to_dict() or {}).get('pin_disposisi_hash') or ''
    if lama_hash and not check_password_hash(lama_hash, pin_lama):
        return jsonify({"error": "PIN lama salah."}), 401
    user_ref.update({'pin_disposisi_hash': generate_password_hash(pin_baru)})
    return jsonify({"message": "PIN disposisi berhasil disimpan."}), 200


# ===================== TELEGRAM LINK/UNLINK ENDPOINTS =====================
@app.route('/api/user/telegram/link', methods=['POST'])
@login_required
def link_telegram():
    """Link user's Telegram account by providing chat_id"""
    user_id = session.get('user_id')
    data = request.get_json(silent=True) or request.form
    chat_id = data.get('chat_id', '').strip()
    
    if not chat_id:
        return jsonify({"error": "Chat ID Telegram wajib diisi."}), 400
    
    # Validate chat_id format (numeric or @username)
    if not (chat_id.startswith('@') or chat_id.lstrip('-').isdigit()):
        return jsonify({"error": "Format Chat ID tidak valid. Gunakan ID numerik atau @username."}), 400
    
    user_ref = fdb.collection('users').document(user_id)
    user_ref.update({'telegram_chat_id': chat_id})
    
    # Send test message
    test_sent = send_telegram_message(chat_id, '✅ <b>Akun Telegram berhasil terhubung!</b>\n\nAnda akan menerima notifikasi tugas baru melalui bot ini.')
    
    if test_sent:
        return jsonify({"message": "Telegram berhasil ditautkan dan pesan tes terkirim."}), 200
    else:
        # Still save but warn
        return jsonify({"message": "Telegram ditautkan, tapi pesan tes gagal. Pastikan bot sudah di-start dan Chat ID benar."}), 200


@app.route('/api/user/telegram/unlink', methods=['POST'])
@login_required
def unlink_telegram():
    """Unlink user's Telegram account"""
    user_id = session.get('user_id')
    user_ref = fdb.collection('users').document(user_id)
    user_ref.update({'telegram_chat_id': DELETE_FIELD})
    
    return jsonify({"message": "Telegram berhasil dilepaskan."}), 200


@app.route('/api/user/telegram/status', methods=['GET'])
@login_required
def telegram_status():
    """Check if user has Telegram linked"""
    user_id = session.get('user_id')
    user_doc = fdb.collection('users').document(user_id).get()
    if not user_doc.exists:
        return jsonify({"error": "Pengguna tidak ditemukan."}), 404
    
    user = user_doc.to_dict()
    return jsonify({
        "linked": bool(user.get('telegram_chat_id')),
        "chat_id": user.get('telegram_chat_id', '')
    }), 200


@app.route('/api/user/delete/<string:user_id>', methods=['DELETE'])
@admin_required
def delete_user(user_id):
    if user_id == session.get('user_id'):
        return jsonify({"error": "Anda tidak dapat menghapus akun Anda sendiri."}), 403

    user_ref = fdb.collection('users').document(user_id)
    if not user_ref.get().exists:
        return jsonify({"error": "Pengguna tidak ditemukan."}), 404

    try:
        user_ref.delete()
        return jsonify({"message": "Pengguna berhasil dihapus."}), 200
    except Exception as e:
        return jsonify({"error": f"Gagal menghapus pengguna: {str(e)}"}), 500

@app.route('/api/user/update/<string:user_id>', methods=['POST'])
@admin_required
def update_user(user_id):
    user_ref = fdb.collection('users').document(user_id)
    if not user_ref.get().exists:
        return jsonify({"error": "Pengguna tidak ditemukan."}), 404

    fullname = request.form.get('fullname')
    role = request.form.get('role')
    jabatan_atasan = (request.form.get('jabatan_atasan') or '').strip()
    bagian = (request.form.get('bagian') or '').strip().lower()
    superior_id_str = request.form.get('superior_id')

    if not fullname or not role:
        return jsonify({"error": "Nama lengkap dan peran harus diisi."}), 400

    if user_id == session.get('user_id') and role != 'admin':
        return jsonify({"error": "Anda tidak dapat mengubah peran Anda sendiri dari admin."}), 403

    # validasi jabatan atasan
    allowed_jabatan = {"dansatsiber","wadansatsiber","asintelops","asminlog","dansat_linta","dansat_indigi","danpus_certan","dansat_dalten",""}
    if jabatan_atasan not in allowed_jabatan:
        jabatan_atasan = ""
    if role != 'atasan':
        jabatan_atasan = ""
    # validasi bagian user
    allowed_bagian = {"persuratan","biasa",""}
    if bagian not in allowed_bagian:
        bagian = ""
    if role != 'user':
        bagian = ""
    # hak disposisi per-pengguna (checkbox admin; tak dicentang = False = ikut bawaan peran)
    disp_kirim = request.form.get('disp_kirim') == '1'
    disp_pantau = request.form.get('disp_pantau') == '1'
    if user_id == session.get('user_id') and role == 'admin':
        disp_kirim, disp_pantau = True, True  # admin tak bisa mencabut haknya sendiri

    try:
        user_ref.update({
            'fullname': fullname,
            'role': role,
            'jabatan_atasan': jabatan_atasan or None,
            'bagian': bagian or None,
            'superior_id': superior_id_str if superior_id_str else None,
            'disp_kirim': disp_kirim,
            'disp_pantau': disp_pantau
        })
        return jsonify({"message": "Data pengguna berhasil diperbarui."}), 200
    except Exception as e:
        return jsonify({"error": f"Gagal memperbarui pengguna: {str(e)}"}), 500


@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('serve_login_page'))

# ==========================================
# ROUTES UNTUK MENYAJIKAN HALAMAN HTML
# ==========================================
@app.route('/dokum/<path:filename>')
@login_required
def serve_dokum(filename):
    """Melayani file scanner iOS & aset statis di folder dokum/ (html, gambar).

    Tanpa route ini, iframe pemindai iOS melempar 404
    'The requested URL was not found on the server'."""
    allowed_ext = {'.html', '.htm', '.png', '.jpg', '.jpeg', '.gif', '.webp', '.svg', '.css', '.js'}
    if os.path.splitext(filename)[1].lower() not in allowed_ext:
        return "Jenis file tidak diizinkan", 403
    dokum_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'dokum')
    return send_from_directory(dokum_dir, filename)

@app.route('/')
def serve_login_page():
    return render_template('LOGIN.html')

@app.route('/register.html')
def serve_register_page():
    return render_template('register.html')

@app.route('/admin/dashboard.html')
@admin_required
def serve_admin_dashboard():
    return render_template('admin/dashboard.html', user_name=session.get('user_name', 'Admin'))

@app.route('/admin/kelola-pengguna.html')
@admin_required
def serve_admin_kelola_pengguna():
    return render_template('admin/kelola-pengguna.html', user_name=session.get('user_name', 'Admin'))

@app.route('/admin/lapor-kendala.html')
@admin_required
def serve_admin_lapor_kendala():
    return render_template('admin/lapor-kendala.html', user_name=session.get('user_name', 'Admin'))

@app.route('/atasan/dashboard.html')
@login_required
def serve_atasan_dashboard():
    return render_template('atasan/dashboard.html', user_name=session.get('user_name', 'Atasan'))

@app.route('/atasan/tim-saya.html')
@login_required
def serve_atasan_tim_saya():
    if session.get('user_role') != 'atasan':
        return "Akses Ditolak", 403
    return render_template('atasan/tim-saya.html', user_name=session.get('user_name', 'Atasan'))

@app.route('/atasan/jadwal-rapat.html')
@login_required
def serve_atasan_jadwal_rapat():
    if session.get('user_role') != 'atasan':
        return "Akses Ditolak", 403
    return render_template('atasan/jadwal-rapat.html', user_name=session.get('user_name', 'Atasan'))

@app.route('/atasan/kalender-kegiatan.html')
@login_required
def serve_atasan_kalender_kegiatan():
    if session.get('user_role') != 'atasan':
        return "Akses Ditolak", 403
    return render_template('atasan/kalender-kegiatan.html', user_name=session.get('user_name', 'Atasan'))

@app.route('/atasan/kelola-pengumuman.html')
@login_required
def serve_atasan_kelola_pengumuman():
    if session.get('user_role') != 'atasan':
        return "Akses Ditolak", 403
    return render_template('atasan/kelola-pengumuman.html', user_name=session.get('user_name', 'Atasan'))

@app.route('/atasan/persetujuan.html')
@login_required
def serve_atasan_persetujuan():
    if session.get('user_role') != 'atasan':
        return "Akses Ditolak", 403
    return render_template('atasan/persetujuan.html', user_name=session.get('user_name', 'Atasan'))

@app.route('/USER/dashboard.html')
@login_required
def serve_user_dashboard():
    return render_template('USER/dashboard.html', user_name=session.get('user_name', 'Pengguna'))

@app.route('/USER/direktori-staf.html')
@login_required
def serve_user_direktori():
    return render_template('USER/direktori-staf.html', user_name=session.get('user_name', 'Pengguna'))

@app.route('/USER/bantuan.html')
@login_required
def serve_user_bantuan():
    return render_template('USER/bantuan.html', user_name=session.get('user_name', 'Pengguna'))

@app.route('/USER/alat-bantuan.html')
@login_required
def serve_user_alat():
    return render_template('USER/alat-bantuan.html', user_name=session.get('user_name', 'Pengguna'))

@app.route('/USER/profile.html')
@login_required
def serve_user_profile():
    return render_template('USER/profile.html', user_name=session.get('user_name', 'Pengguna'), user_email=session.get('user_email', ''))

@app.route('/USER/riwayat-disposisi.html')
@login_required
@disposisi_required
def serve_user_riwayat_disposisi():
    return render_template('USER/riwayat-disposisi.html', user_name=session.get('user_name', 'Pengguna'))

@app.route('/USER/jadwal-rapat.html')
@login_required
def serve_user_jadwal_rapat():
    return render_template('USER/jadwal-rapat.html', user_name=session.get('user_name', 'Pengguna'))

@app.route('/USER/kalender-kegiatan.html')
@login_required
def serve_user_kalender_kegiatan():
    return render_template('USER/kalender-kegiatan.html', user_name=session.get('user_name', 'Pengguna'))

@app.route('/admin/jadwal-rapat.html')
@admin_required
def serve_admin_jadwal_rapat():
    return render_template('admin/jadwal-rapat.html', user_name=session.get('user_name', 'Admin'))

@app.route('/admin/kalender-kegiatan.html')
@admin_required
def serve_admin_kalender_kegiatan():
    return render_template('admin/kalender-kegiatan.html', user_name=session.get('user_name', 'Admin'))

@app.route('/admin/kelola-pengumuman.html')
@admin_required
def serve_admin_kelola_pengumuman():

   return render_template('admin/kelola-pengumuman.html', user_name=session.get('user_name', 'Admin'))

@app.route('/api/pengumuman', methods=['GET'])
@login_required
def get_announcements():
    docs = fdb.collection('pengumuman').stream()
    hasil = []
    for d in docs:
        a = d.to_dict()
        hasil.append({
            'id': d.id,
            'title': a.get('title'),
            'content': a.get('content'),
            'category': a.get('category', 'indigo'),
            'date': a.get('date'),
            'created_date': a.get('created_date'),
            'created_time': a.get('created_time'),
            'publish_date': a.get('publish_date'),
            'publish_time': a.get('publish_time')
        })
    hasil.reverse()  # terbaru dulu
    return jsonify(hasil)

@app.route('/api/pengumuman', methods=['POST'])
@login_required
def add_announcement():
    data = request.get_json(silent=True) or {}
    title = data.get('title')
    content = data.get('content')
    category = data.get('category', 'indigo')

    if not title or not content:
        return jsonify({"error": "Judul dan konten wajib diisi."}), 400

    try:
        fdb.collection('pengumuman').document().set({
            'title': title,
            'content': content,
            'category': category,
            'date': datetime.datetime.now().strftime('%d %b %Y %H:%M'),
            'created_date': data.get('created_date'),
            'created_time': data.get('created_time'),
            'publish_date': data.get('publish_date'),
            'publish_time': data.get('publish_time')
        })
        return jsonify({"message": "Pengumuman berhasil ditambahkan!"}), 201
    except Exception as e:
        return jsonify({"error": f"Gagal menambah pengumuman: {str(e)}"}), 500

@app.route('/api/pengumuman/<string:ann_id>', methods=['DELETE'])
@login_required
def delete_announcement(ann_id):
    doc_ref = fdb.collection('pengumuman').document(ann_id)
    if not doc_ref.get().exists:
        return jsonify({"error": "Pengumuman tidak ditemukan."}), 404

    try:
        doc_ref.delete()
        return jsonify({"message": "Pengumuman berhasil dihapus!"}), 200
    except Exception as e:
        return jsonify({"error": f"Gagal menghapus pengumuman: {str(e)}"}), 500

@app.route('/api/meetings', methods=['GET'])
@login_required
def get_meetings():
    docs = fdb.collection('jadwal_rapat').stream()
    hasil = []
    for d in docs:
        m = d.to_dict()
        raw = m.get('tanggal') or ''
        try:
            pretty = datetime.datetime.strptime(raw, '%Y-%m-%d').strftime('%d %b %Y') if raw else ''
        except Exception:
            pretty = raw
        hasil.append({
            'id': d.id,
            'judul': m.get('judul'),
            'judul_rapat': m.get('judul'),
            'tanggal': pretty,
            'tanggal_raw': raw,
            'tanggal_rapat': raw,
            'waktu': m.get('waktu'),
            'waktu_rapat': m.get('waktu'),
            'peserta': m.get('peserta'),
            'peserta_rapat': m.get('peserta')
        })
    hasil.sort(key=lambda m: (m['tanggal_raw'] or '', m['waktu'] or ''))
    return jsonify(hasil)

@app.route('/api/meetings', methods=['POST'])
@admin_required
def add_meeting():
    judul = request.form.get('judul_rapat')
    tanggal_str = request.form.get('tanggal_rapat')
    waktu_str = request.form.get('waktu_rapat')
    peserta = request.form.get('peserta_rapat')

    if not all([judul, tanggal_str, waktu_str]):
        return jsonify({"error": "Judul, tanggal, dan waktu harus diisi."}), 400

    try:
        # Validasi format saja, disimpan tetap sebagai string 'YYYY-MM-DD' dan 'HH:MM'
        datetime.datetime.strptime(tanggal_str, '%Y-%m-%d')
        datetime.datetime.strptime(waktu_str, '%H:%M')

        fdb.collection('jadwal_rapat').document().set({
            'judul': judul,
            'tanggal': tanggal_str,
            'waktu': waktu_str,
            'peserta': peserta
        })

        return jsonify({"message": "Jadwal rapat berhasil ditambahkan."}), 201
    except Exception as e:
        return jsonify({"error": f"Gagal menambah jadwal: {str(e)}"}), 500

@app.route('/api/meetings/<string:meeting_id>', methods=['DELETE'])
@admin_required
def delete_meeting(meeting_id):
    doc_ref = fdb.collection('jadwal_rapat').document(meeting_id)
    if not doc_ref.get().exists:
        return jsonify({"error": "Jadwal rapat tidak ditemukan."}), 404
    try:
        doc_ref.delete()
        return jsonify({"message": "Jadwal rapat berhasil dihapus."}), 200
    except Exception as e:
        return jsonify({"error": f"Gagal menghapus jadwal: {str(e)}"}), 500


# ===================== PUSAT BANTUAN: TIKET USER -> ADMIN =====================
# Alur: user kirim tiket (kategori/subjek/pesan) -> admin lihat di inbox
# dashboard -> admin balas -> user baca balasan di dashboard/bantuan.
def _serialize_tiket(doc, pelapor=None):
    t = doc.to_dict() or {}
    u = pelapor or {}
    return {
        'id': doc.id,
        'pelapor_id': t.get('user_id') or '',
        'pelapor_nama': t.get('pelapor_nama') or '-',
        'pelapor_nrp': u.get('nrp') or '',
        'pelapor_email': u.get('email') or '',
        'pelapor_bagian': u.get('bagian') or '',
        'kategori': t.get('kategori') or '-',
        'subjek': t.get('subjek') or '-',
        'pesan': t.get('pesan') or '',
        'status': t.get('status') or 'terbuka',
        'balasan': t.get('balasan') or '',
        'dibalas_pada': t.get('dibalas_pada') or '',
        'dibaca_user': bool(t.get('dibaca_user', True)),
        'created_at': t.get('created_at') or ''
    }


@app.route('/api/bantuan/kirim', methods=['POST'])
@login_required
def kirim_tiket_bantuan():
    try:
        data = request.get_json(force=True, silent=True) or {}
        kategori = str(data.get('kategori') or '').strip()[:60]
        subjek = str(data.get('subjek') or '').strip()[:200]
        pesan = str(data.get('pesan') or '').strip()[:2000]
        if not subjek or not pesan:
            return jsonify({"error": "Subjek dan pesan wajib diisi."}), 400
        saya = session.get('user_id')
        nama = session.get('user_name') or _nama_pengguna(saya)
        doc_ref = fdb.collection('bantuan_tiket').document()
        doc_ref.set({
            'user_id': saya,
            'pelapor_nama': nama,
            'kategori': kategori or 'Lainnya',
            'subjek': subjek,
            'pesan': pesan,
            'status': 'terbuka',
            'balasan': '',
            'dibalas_pada': '',
            'dibaca_user': True,
            'created_at': datetime.datetime.now().strftime('%d %b %Y %H:%M')
        })
        return jsonify({"message": "Tiket bantuan terkirim ke admin.", "id": doc_ref.id}), 201
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route('/api/bantuan/milik-saya', methods=['GET'])
@login_required
def tiket_milik_saya():
    try:
        saya = session.get('user_id')
        docs = fdb.collection('bantuan_tiket').where('user_id', '==', saya).stream()
        hasil = [_serialize_tiket(d) for d in docs]
        hasil.sort(key=lambda h: h['created_at'], reverse=True)
        return jsonify(hasil), 200
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route('/api/bantuan/semua', methods=['GET'])
@admin_required
def semua_tiket_bantuan():
    try:
        # Ambil data pelapor sekali agar admin bisa menghubungi user langsung.
        pengguna = {}
        for u in fdb.collection('users').stream():
            pengguna[u.id] = u.to_dict() or {}
        docs = fdb.collection('bantuan_tiket').stream()
        hasil = [_serialize_tiket(d, pengguna.get((d.to_dict() or {}).get('user_id'))) for d in docs]
        hasil.sort(key=lambda h: h['created_at'], reverse=True)
        return jsonify(hasil), 200
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route('/api/bantuan/<string:tiket_id>/balas', methods=['POST'])
@admin_required
def balas_tiket_bantuan(tiket_id):
    try:
        data = request.get_json(force=True, silent=True) or {}
        balasan = str(data.get('balasan') or '').strip()[:2000]
        if not balasan:
            return jsonify({"error": "Balasan wajib diisi."}), 400
        doc_ref = fdb.collection('bantuan_tiket').document(tiket_id)
        if not doc_ref.get().exists:
            return jsonify({"error": "Tiket tidak ditemukan."}), 404
        doc_ref.update({
            'balasan': balasan,
            'dibalas_pada': datetime.datetime.now().strftime('%d %b %Y %H:%M'),
            'status': 'dibalas',
            'dibaca_user': False
        })
        return jsonify({"message": "Balasan terkirim ke pelapor."}), 200
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route('/api/bantuan/<string:tiket_id>/baca', methods=['POST'])
@login_required
def tandai_tiket_dibaca(tiket_id):
    try:
        doc_ref = fdb.collection('bantuan_tiket').document(tiket_id)
        doc = doc_ref.get()
        if not doc.exists:
            return jsonify({"error": "Tiket tidak ditemukan."}), 404
        t = doc.to_dict() or {}
        if t.get('user_id') != session.get('user_id') and session.get('user_role') != 'admin':
            return jsonify({"error": "Bukan tiket Anda."}), 403
        doc_ref.update({'dibaca_user': True})
        return jsonify({"message": "OK"}), 200
    except Exception as e:
        return jsonify({"error": str(e)}), 500


# ===================== PDF Panduan Portal =====================
@app.route('/api/panduan-portal', methods=['GET'])
@login_required
def download_panduan_portal():
    """Generate and serve PDF user guide for the portal."""
    try:
        from docx import Document
        from docx.shared import Inches, Pt, Cm, RGBColor
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        import io
        
        doc = Document()
        
        # Style setup
        style = doc.styles['Normal']
        font = style.font
        font.name = 'Calibri'
        font.size = Pt(11)
        style.paragraph_format.space_after = Pt(6)
        style.paragraph_format.line_spacing = 1.15
        
        # Title
        title = doc.add_heading('BUKU PANDUAN PENGGUNA PORTAL KARYAWAN', level=0)
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for run in title.runs:
            run.font.color.rgb = RGBColor(0x07, 0x59, 0x85)  # corporate-700
            run.font.size = Pt(24)
        
        subtitle = doc.add_paragraph('Sistem Informasi Kantor - Versi 3.2.0')
        subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
        subtitle.runs[0].font.size = Pt(12)
        subtitle.runs[0].font.color.rgb = RGBColor(0x64, 0x74, 0x8B)  # slate-500
        
        doc.add_paragraph('')  # spacer
        
        # Table of Contents
        doc.add_heading('DAFTAR ISI', level=1)
        toc_items = [
            '1. Pendahuluan, Peran Pengguna & Login',
            '2. Dashboard (Beranda Utama)',
            '3. Tugas Saya - Manajemen Tugas Pribadi & dari Atasan',
            '4. Direktori Staf',
            '5. Alat Bantu Dokumen (Konversi, Scan & Ringkasan)',
            '6. Disposisi Surat - Kirim, Terima & Teruskan',
            '7. Jadwal Rapat',
            '8. Kalender Kegiatan',
            '9. Profil Saya & Pengaturan Akun',
            '10. Pusat Bantuan, Tiket & Balasan Admin',
            '11. Tips & Trik Produktivitas'
        ]
        for item in toc_items:
            p = doc.add_paragraph(item)
            p.paragraph_format.space_after = Pt(2)
        
        doc.add_page_break()
        
        # Chapter 1
        doc.add_heading('1. PENDAHULUAN, PERAN PENGGUNA & LOGIN', level=1)
        doc.add_paragraph(
            'Portal Karyawan adalah sistem informasi terintegrasi untuk mengelola aktivitas kerja harian, '
            'tugas, jadwal rapat, dan komunikasi internal kantor. Sistem ini dirancang dengan antarmuka '
            'modern yang responsif dan mudah digunakan di berbagai perangkat (desktop, tablet, mobile).'
        )
        doc.add_paragraph('Fitur utama:')
        features = [
            'Manajemen tugas real-time (diberikan atasan & pribadi)',
            'Kalender kegiatan & jadwal rapat terintegrasi',
            'Disposisi surat berjenjang: kirim, terima, teruskan, dan unduh PDF',
            'Konversi dokumen (PDF ↔ Word), scan dokumen, dan peringkasan otomatis',
            'Direktori staf & struktur organisasi',
            'Pengumuman kedinasan berwarna',
            'Pusat Bantuan dengan tiket kendala dan balasan admin',
            'Profil pengguna dengan foto & biodata lengkap'
        ]
        for f in features:
            doc.add_paragraph(f, style='List Bullet')
        
        doc.add_heading('Peran Pengguna', level=2)
        doc.add_paragraph(
            'Setiap akun memiliki peran yang menentukan menu dan hak akses:'
        )
        peran = [
            'Admin - mengelola pengguna, struktur organisasi, pengumuman, jadwal rapat, '
            'format surat, serta membalas tiket bantuan. Admin memiliki akses penuh ke disposisi.',
            'Atasan - memberi tugas, menyetujui atau menolak hasil kerja staf, menerima surat '
            'disposisi, mengonfirmasi penerimaan, dan meneruskan disposisi ke bawahan, '
            'staf/persuratan, atau atasan lain.',
            'User (Staf) - mengerjakan tugas, melihat direktori staf, memakai Alat Bantu Dokumen, '
            'mengirim tiket bantuan, dan melihat riwayat disposisi bila memiliki hak persuratan.'
        ]
        for p in peran:
            doc.add_paragraph(p, style='List Bullet')
        
        doc.add_heading('Cara Login', level=2)
        login_steps = [
            'Buka browser dan akses alamat portal yang diberikan Admin (contoh: https://portal.domain.com)',
            'Masukkan email perusahaan Anda',
            'Masukkan kata sandi',
            'Klik tombol "Masuk"',
            'Sistem akan mengarahkan ke dashboard sesuai peran (Admin/Atasan/Staf)'
        ]
        for i, step in enumerate(login_steps, 1):
            doc.add_paragraph(f'{step}', style='List Number')
        
        doc.add_paragraph(
            'Catatan: Jika lupa kata sandi, hubungi Helpdesk IT/Admin untuk dilakukan reset manual.'
        )
        
        # Chapter 2
        doc.add_heading('2. DASHBOARD (BERANDA UTAMA)', level=1)
        doc.add_paragraph(
            'Dashboard adalah halaman pertama yang muncul setelah login. Menampilkan ringkasan:'
        )
        dash_items = [
            'Statistik Tugas: Aktif, Selesai, Pengumuman',
            'Daftar 5 tugas terbaru yang diberikan atasan',
            'Pengumuman kedinasan terbaru (maksimal 5)',
            'Navigasi cepat ke fitur-fitur utama via sidebar'
        ]
        for item in dash_items:
            doc.add_paragraph(item, style='List Bullet')
        
        doc.add_paragraph(
            'Klik "Lihat Semua" pada bagian Tugas Terbaru untuk membuka halaman "Tugas Saya" lengkap dengan filter status.'
        )
        
        # Chapter 3
        doc.add_heading('3. TUGAS SAYA - MANAJEMEN TUGAS PRIBADI & DARI ATASAN', level=1)
        doc.add_paragraph(
            'Halaman ini menampilkan semua tugas Anda dalam tabel interaktif dengan fitur:'
        )
        task_features = [
            'Filter status: Semua / Belum Dikerjakan / Sedang Dikerjakan / Selesai',
            'Form tambah tugas pribadi (judul, deskripsi, prioritas, deadline)',
            'Edit & hapus tugas pribadi (tugas dari atasan tidak bisa diedit/dihapus)',
            'Ubah status tugas via dropdown (real-time update)',
            'Tampilkan nama pemberi tugas (atasan)',
            'Pembaruan otomatis tiap 10 detik (tanpa refresh manual)'
        ]
        for f in task_features:
            doc.add_paragraph(f, style='List Bullet')
        
        doc.add_heading('Prioritas Tugas', level=2)
        doc.add_paragraph('Tinggi (Merah) - Mendesak, harus diselesaikan segera')
        doc.add_paragraph('Sedang (Kuning) - Biasa, sesuai jadwal')
        doc.add_paragraph('Rendah (Abu-abu) - Fleksibel, tidak terburu-buru')
        
        # Chapter 4
        doc.add_heading('4. DIREKTORI STAF', level=1)
        doc.add_paragraph(
            'Menampilkan daftar seluruh karyawan dengan informasi: nama, email, peran (Admin/Atasan/Staf), '
            'dan atasan langsung. Dilengkapi pencarian real-time.'
        )
        
        # Chapter 5
        doc.add_heading('5. ALAT BANTU DOKUMEN', level=1)
        doc.add_paragraph(
            'Menu "Alat Bantu Dokumen" menyediakan beberapa fitur untuk mengolah berkas kantor:'
        )
        tools = [
            'PDF ke Word (.docx) - mengubah PDF menjadi dokumen Word yang bisa diedit. '
            'PDF hasil scan diproses dengan OCR agar teksnya dapat dikenali.',
            'Word ke PDF - mengonversi .docx menjadi PDF dengan layout tetap. '
            'Fitur ini membutuhkan Microsoft Word yang terpasang di server.',
            'Ringkas Dokumen - mengunggah .docx lalu merangkum isinya otomatis '
            'menggunakan NLP (LexRank + Sastrawi untuk Bahasa Indonesia).',
            'Jelaskan Dokumen - menampilkan topik utama dan poin penting dari dokumen.',
            'Scan Dokumen - memindai dokumen melalui scanner USB di komputer pengguna '
            '(Scan Agent lokal) atau scanner yang terhubung ke server (WIA) menjadi PDF '
            'multi-halaman. Mendukung scanner flatbed seperti Epson L3210 dan ADF/duplex '
            'seperti Epson DS-410.',
            'Panduan Portal - mengunduh Buku Panduan Pengguna dalam format DOCX.'
        ]
        for t in tools:
            doc.add_paragraph(t, style='List Bullet')
        
        doc.add_heading('Langkah Menggunakan Konversi', level=2)
        konversi_steps = [
            'Buka menu "Alat Bantu Dokumen".',
            'Pilih jenis konversi yang diinginkan (PDF ke Word atau Word ke PDF).',
            'Klik area upload lalu pilih berkas dari komputer.',
            'Tunggu proses selesai; hasil akan terunduh otomatis.',
            'Jika browser memblokir unduhan, izinkan popup/unduhan untuk portal ini.'
        ]
        for s in konversi_steps:
            doc.add_paragraph(s, style='List Number')
        
        doc.add_heading('Langkah Menggunakan Scan Dokumen', level=2)
        scan_steps = [
            'Pastikan scanner menyala dan terpasang di komputer Anda atau di server.',
            'Jika scanner USB terpasang di komputer Anda, jalankan Scan Agent lokal '
            '(klik dua kali JALANKAN-SCAN-AGENT.bat) dan biarkan jendelanya terbuka.',
            'Buka menu "Alat Bantu Dokumen" lalu pilih bagian Scan Dokumen.',
            'Pilih perangkat scanner, mode warna, resolusi, dan opsi duplex bila memakai ADF.',
            'Letakkan dokumen di scanner, lalu klik tombol Scan.',
            'Periksa hasil scan, lalu unduh atau lampirkan ke disposisi.'
        ]
        for s in scan_steps:
            doc.add_paragraph(s, style='List Number')
        
        doc.add_paragraph(
            'Catatan: Portal otomatis mencoba scanner di komputer Anda lebih dulu, lalu '
            'scanner yang terpasang di server. Jika scanner tidak terdeteksi, pastikan Scan '
            'Agent lokal berjalan atau hubungi Helpdesk. Anda juga dapat mengunggah PDF hasil '
            'scan manual.'
        )
        
        # Chapter 6 - Disposisi Surat
        doc.add_heading('6. DISPOSISI SURAT - KIRIM, TERIMA & TERUSKAN', level=1)
        doc.add_paragraph(
            'Fitur Disposisi Surat digunakan untuk mengirim surat dinas secara berjenjang, '
            'mencatat siapa yang menerima, dan meneruskan surat kepada pihak lain. '
            'Seluruh jejak disposisi tersimpan di sistem.'
        )
        
        doc.add_heading('6.1 Mengirim Surat Disposisi', level=2)
        kirim_disposisi = [
            'Buka menu "Alat Bantu Dokumen" lalu pilih form Disposisi Surat '
            '(pengguna persuratan) atau menu "Riwayat Disposisi".',
            'Isi identitas surat: nomor agenda, asal surat, nomor surat, tanggal, dan perihal.',
            'Pilih format lembar disposisi yang tersedia.',
            'Tentukan tujuan "Diteruskan Kepada" dan centang isi disposisi yang sesuai.',
            'Tambahkan catatan dan paraf bila diperlukan.',
            'Unggah berkas surat (PDF) dan lampiran bila ada, atau gunakan hasil Scan Dokumen.',
            'Klik Preview untuk memeriksa hasil isian, lalu Unduh PDF bila perlu.',
            'Klik Kirim untuk mengirim disposisi kepada atasan/tujuan yang dipilih.'
        ]
        for s in kirim_disposisi:
            doc.add_paragraph(s, style='List Number')
        
        doc.add_heading('6.2 Menerima dan Meneruskan Surat', level=2)
        doc.add_paragraph(
            'Atasan menerima surat pada halaman "Persetujuan". Langkah yang dapat dilakukan:'
        )
        terima_disposisi = [
            'Buka menu Persetujuan lalu pilih tab surat disposisi yang masuk.',
            'Klik Lihat Surat untuk membaca isi disposisi dan lampiran.',
            'Klik Konfirmasi Diterima untuk mencatat bahwa surat sudah diterima.',
            'Klik Teruskan untuk meneruskan surat kepada bawahan langsung, staf/persuratan, '
            'atau atasan lain. Tujuan dapat dipilih lebih dari satu.',
            'Isi format disposisi, centang "Diteruskan Kepada" dan "Isi Disposisi", '
            'tambahkan catatan serta paraf.',
            'Gunakan Preview Hasil untuk melihat hasil isian Anda sebelum dikirim.',
            'Klik Unduh PDF untuk mengunduh lembar disposisi, atau Teruskan untuk mengirim.'
        ]
        for s in terima_disposisi:
            doc.add_paragraph(s, style='List Number')
        
        doc.add_paragraph(
            'Penting: preview dan PDF pada proses Teruskan dibuat dari isian terbaru Anda, '
            'bukan menyalin surat asal. Penerima akan melihat disposisi sesuai isian yang Anda kirim.'
        )
        
        doc.add_heading('6.3 Status dan Jejak Surat', level=2)
        status_disposisi = [
            'Baru - surat sudah dikirim dan menunggu dibaca/diproses penerima.',
            'Dikonfirmasi - penerima sudah menyatakan surat diterima.',
            'Jejak "Diteruskan" pada riwayat mencatat ke mana surat diteruskan dan oleh siapa.',
            'Riwayat Disposisi menampilkan surat yang pernah Anda kirim/terima beserta rantai penerusannya.'
        ]
        for s in status_disposisi:
            doc.add_paragraph(s, style='List Bullet')
        
        doc.add_heading('6.4 Hak Akses Disposisi', level=2)
        doc.add_paragraph(
            'Hak mengirim dan meneruskan disposisi mengikuti peran dan pengaturan Admin. '
            'Sebagian pengguna hanya dapat memantau jejak surat, sedangkan pengguna persuratan '
            'dan atasan tertentu dapat mengirim atau meneruskan. Jika tombol tidak muncul, '
            'hubungi Admin untuk memeriksa hak akses akun Anda.'
        )
        
        # Chapter 7
        doc.add_heading('7. JADWAL RAPAT', level=1)
        doc.add_paragraph(
            'Kelola jadwal rapat kantor. Admin bisa menambah/hapus jadwal. Semua pengguna bisa melihat jadwal.'
        )
        meeting_features = [
            'Form tambah rapat: Judul, Tanggal, Waktu, Peserta',
            'Daftar rapat terurut kronologis',
            'Hapus jadwal (hanya Admin)',
            'Tampilan terintegrasi di Kalender Kegiatan'
        ]
        for f in meeting_features:
            doc.add_paragraph(f, style='List Bullet')
        
        # Chapter 8
        doc.add_heading('8. KALENDER KEGIATAN', level=1)
        doc.add_paragraph(
            'Kalender bulanan interaktif menampilkan:'
        )
        cal_items = [
            'Pengumuman (berdasarkan tanggal & waktu pelaksanaan)',
            'Jadwal Rapat (berdasarkan tanggal & waktu rapat)',
            'Hari Libur Nasional 2026 (merah)',
            'Hari ini (highlight biru)',
            'Navigasi bulan Prev/Next/Today',
            'Maksimal 3 event per hari, keterangan "+N lagi..." jika lebih'
        ]
        for item in cal_items:
            doc.add_paragraph(item, style='List Bullet')
        
        doc.add_paragraph(
            'Catatan: Tanggal pelaksanaan pengumuman diatur di "Kelola Pengumuman" (Admin). '
            'Jadwal rapat diatur di "Jadwal Rapat" (Admin).'
        )
        
        # Chapter 9
        doc.add_heading('9. PROFIL SAYA & PENGATURAN AKUN', level=1)
        doc.add_paragraph('Dua tab utama:')
        
        doc.add_heading('Biodata Pribadi', level=2)
        bio_fields = [
            'Foto Profil (klik ikon kamera untuk upload/ganti)',
            'Nama Lengkap, NRP (angka), Asal, Gender (Laki-laki/Perempuan)',
            'Email (@gmail.com), Nomor Telepon (angka)',
            'Data tersimpan permanen di MySQL, tetap ada setelah logout/login'
        ]
        for f in bio_fields:
            doc.add_paragraph(f, style='List Bullet')
        
        doc.add_heading('Ubah Kata Sandi', level=2)
        doc.add_paragraph('Masukkan kata sandi lama, kata sandi baru (min. 4 karakter), dan konfirmasi.')
        
        # Chapter 10
        doc.add_heading('10. PUSAT BANTUAN, TIKET & BALASAN ADMIN', level=1)
        doc.add_paragraph(
            'Pusat Bantuan adalah tempat untuk membaca panduan, melihat pertanyaan yang sering '
            'diajukan (FAQ), dan mengirim tiket kendala kepada Admin/Helpdesk IT.'
        )
        
        doc.add_heading('10.1 Mengirim Tiket Bantuan', level=2)
        tiket_steps = [
            'Buka menu "Pusat Bantuan" pada sidebar.',
            'Pilih kategori kendala: Kendala IT & Akses Sistem, Pertanyaan Administrasi, '
            'Sarana & Fasilitas Kantor, atau Lainnya.',
            'Isi subjek permasalahan secara singkat dan jelas.',
            'Tulis deskripsi lengkap: kronologi, waktu kejadian, pesan error, dan dampaknya.',
            'Klik "Kirim Tiket Bantuan".',
            'Tiket akan muncul pada bagian "Tiket Saya & Balasan Admin" dengan status menunggu.'
        ]
        for s in tiket_steps:
            doc.add_paragraph(s, style='List Number')
        
        doc.add_heading('10.2 Melihat Balasan Admin', level=2)
        balasan_steps = [
            'Buka kembali menu "Pusat Bantuan".',
            'Lihat bagian "Tiket Saya & Balasan Admin".',
            'Tiket yang sudah dibalas akan berstatus "dibalas admin" dan ditandai hijau.',
            'Baca balasan pada kotak "Balasan Admin" di bawah deskripsi tiket.',
            'Jika masalah belum selesai, kirim tiket baru atau balas melalui koordinasi dengan Admin.'
        ]
        for s in balasan_steps:
            doc.add_paragraph(s, style='List Number')
        
        doc.add_heading('10.3 Proses di Sisi Admin', level=2)
        doc.add_paragraph(
            'Tiket yang dikirim user otomatis masuk ke dashboard Admin. Admin dapat melihatnya '
            'pada "Inbox Laporan Pusat Bantuan" di dashboard atau halaman "Laporan Kendala". '
            'Admin dapat mencari tiket, memfilter status, membaca detail pelapor, dan mengirim balasan. '
            'Setelah dibalas, status tiket berubah menjadi "dibalas" dan balasan tampil di halaman user.'
        )
        
        doc.add_heading('10.4 Buku Panduan Pengguna', level=2)
        doc.add_paragraph(
            'Pada halaman Pusat Bantuan tersedia kartu "Buku Panduan Pengguna". '
            'Klik "Unduh Panduan (DOCX)" untuk mengunduh dokumen panduan ini. '
            'Pastikan Anda sudah login agar berkas dapat diunduh.'
        )
        
        doc.add_heading('10.5 Kontak Helpdesk', level=2)
        doc.add_paragraph(
            'Nomor hotline dan email layanan support ditampilkan pada kartu bagian atas '
            'halaman Pusat Bantuan. Gunakan kontak tersebut untuk kendala yang mendesak.'
        )
        
        # Chapter 11
        doc.add_heading('11. TIPS & TRIK PRODUKTIVITAS', level=1)
        tips = [
            'Gunakan filter status di "Tugas Saya" untuk fokus pada tugas "Sedang Dikerjakan"',
            'Atur deadline realistis agar terlihat di Kalender Kegiatan',
            'Upload foto profil agar mudah dikenali di Direktori Staf',
            'Manfaatkan "Ringkas Dokumen" untuk membaca laporan panjang secara singkat',
            'Cek Kalender Kegiatan mingguan untuk perencanaan jadwal',
            'Gunakan mode mobile (hamburger menu) di smartphone untuk akses cepat',
            'Selalu gunakan "Preview Hasil" sebelum meneruskan surat disposisi',
            'Klik "Konfirmasi Diterima" setelah surat disposisi benar-benar diterima',
            'Periksa "Tiket Saya & Balasan Admin" secara berkala agar tidak melewatkan balasan',
            'Untuk scan dokumen, gunakan resolusi 200-300 dpi agar hasil jelas namun tidak terlalu besar',
            'Logout hanya saat diperlukan - sesi tetap aktif untuk kemudahan akses'
        ]
        for tip in tips:
            doc.add_paragraph(tip, style='List Bullet')
        
        # Footer
        doc.add_paragraph('')
        doc.add_paragraph('')
        footer = doc.add_paragraph('—')
        footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
        footer = doc.add_paragraph('Portal Karyawan - Sistem Informasi Kantor')
        footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
        footer.runs[0].font.size = Pt(10)
        footer.runs[0].font.color.rgb = RGBColor(0x94, 0xA3, 0xB8)
        footer = doc.add_paragraph('Versi 3.2.0 | Dikembangkan untuk kebutuhan internal')
        footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
        footer.runs[0].font.size = Pt(9)
        footer.runs[0].font.color.rgb = RGBColor(0x94, 0xA3, 0xB8)
        
        # Save to BytesIO
        buffer = io.BytesIO()
        doc.save(buffer)
        buffer.seek(0)
        
        from flask import send_file
        return send_file(
            buffer,
            as_attachment=True,
            download_name='Panduan_Portal_Karyawan.docx',
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )
        
    except Exception as e:
        return jsonify({"error": f"Gagal generate panduan: {str(e)}"}), 500


# ===================== FORMAT SURAT DISPOSISI — DINAMIS (dokum/ + database) =====================
# Format bawaan (4 file asli di folder dokum/). Format tambahan yang dibuat admin
# via halaman Kelola Format Surat disimpan sebagai file .docx di folder dokum/
# + metadata di tabel MySQL 'format_surat', lalu otomatis muncul di opsi
# user & atasan.
DISPOSISI_VARIANTS_BASE = {
    "dansat":  {"label": "Dansatsiber (Lembar_Disposisi_Satsiber_TNI.docx)", "kepada": "YTH. DANSATSIBER TNI", "terus": ["WADANSATSIBER TNI","ASINTELOPS","ASMINLOG","DANPUSCERTAN","DANSATLINTA","DANSATINDIGI","DANSATDALTEN","PABANDYA UM","PABANDYA PERS","PABANDYA LOG","SPRI","PABANDA KU"], "filename": "Lembar_Disposisi_Satsiber_TNI.docx"},
    "wadan":   {"label": "Wadansatsiber (Wadansatsiber.docx)", "kepada": "YTH. WADANSATSIBER TNI", "terus": ["ASINTELOPS","ASMINLOG","DANPUSCERTAN","DANSATLINTA","DANSATINDIGI","DANSATDALTEN","PABANDYA UM","PABANDYA PERS","PABANDYA LOG","SPRI","PABANDA KU"], "filename": "Wadansatsiber.docx"},
    "asintelops": {"label": "Asintelops (asintelops.docx)", "kepada": "YTH. ASINTELOPS SATSIBER TNI", "terus": ["DANPUSCERTAN","DANSATLINTA","DANSATINDIGI","DANSATDALTEN","PABAN INTEL","PABAN OPSLAT","PABAN RENPROGRAR","STAF INTELOPS"], "filename": "asintelops.docx"},
    "asminlog": {"label": "Asminlog (ASIMINLOG.docx)", "kepada": "YTH. ASMINLOG SATSIBER TNI", "terus": ["ASINTELOPS","DANPUSCERTAN","DANSATLINTA","DANSATINDIGI","DANSATDALTEN","PABANDYA UM","PABANDYA PERS","PABANDYA LOG","SPRI","PABANDA KU"], "filename": "ASIMINLOG.docx"},
}
# Alias lama agar kode lain tetap jalan
DISPOSISI_VARIANTS = DISPOSISI_VARIANTS_BASE

def _dokum_dir():
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), "dokum")

def _slugify_key(nama):
    import re
    s = (nama or '').strip().lower()
    s = re.sub(r'[^a-z0-9]+', '_', s).strip('_')
    return s or ('format_' + datetime.datetime.now().strftime('%Y%m%d%H%M%S'))

def get_all_variants():
    """Gabung format bawaan + format buatan admin (database) + file .docx baru di dokum/."""
    hasil = {}
    for k, v in DISPOSISI_VARIANTS_BASE.items():
        hasil[k] = dict(v)
    # 1) metadata dari database
    try:
        if fdb is not None:
            for d in fdb.collection('format_surat').stream():
                m = d.to_dict() or {}
                key = (m.get('key') or d.id).strip().lower()
                if not key or key in hasil:
                    continue
                hasil[key] = {
                    'label': m.get('label') or key,
                    'kepada': m.get('kepada') or '',
                    'terus': m.get('terus') or [],
                    'filename': m.get('filename') or '',
                    'custom': True,
                    'doc_id': d.id,
                }
    except Exception as e:
        print(f"[FORMAT] gagal baca format_surat dari database: {e}")
    # 2) file .docx di dokum/ yang belum terdaftar -> daftarkan otomatis
    try:
        dokum = _dokum_dir()
        dikenal = { (v.get('filename') or '').lower() for v in hasil.values() }
        for fn in sorted(os.listdir(dokum)):
            if not fn.lower().endswith('.docx') or fn.startswith('~'):
                continue
            if fn.lower() in dikenal:
                continue
            key = _slugify_key(os.path.splitext(fn)[0])[:40]
            if key in hasil:
                key = key + '_' + datetime.datetime.now().strftime('%H%M%S')
            hasil[key] = {
                'label': f"{os.path.splitext(fn)[0]} ({fn})",
                'kepada': f"YTH. {os.path.splitext(fn)[0].upper()}",
                'terus': list(DISPOSISI_VARIANTS_BASE['dansat']['terus']),
                'filename': fn,
                'custom': True,
                'auto': True,
            }
    except Exception as e:
        print(f"[FORMAT] gagal scan dokum/: {e}")
    return hasil

def get_variant_info(varian_key):
    semua = get_all_variants()
    v = (varian_key or 'dansat').strip().lower()
    if v not in semua:
        v = 'dansat'
    return v, semua[v], semua
# backward compat
DISPOSISI_DITERUSKAN = DISPOSISI_VARIANTS_BASE["dansat"]["terus"]
DISPOSISI_ISI_1 = ["UDK","UDL","UMP","CATAT","ACC","WAKILI","PELAJARI/TELITI","KOORDINASIKAN","FASILITASI","PEDOMANI","EDARKAN/INFOKAN JAJARAN","DUKUNG","IKUTI BANGSIT"]
DISPOSISI_ISI_2 = ["JADWALKAN","TINDAKLANJUTI","SIAPKAN PERSONEL","PAPARKAN","SIAPKAN JAWABAN","SIAPKAN BAHAN","MONITOR","COPY","SARAN","INGATKAN","LAPORKAN HASILNYA","PERBAHARUI DATABASE","ARSIP"]

def _build_disposisi_docx(data):
    from docx import Document
    from docx.shared import Pt, Inches
    from docx.oxml.ns import qn
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    import io, os, base64, datetime
    varian, var_info, _semua = get_variant_info(data.get('varian') or data.get('template') or 'dansat')
    kepada_text = var_info.get("kepada") or ""
    kepada_custom = (data.get('kepada_custom') or data.get('kepada_nama') or '').strip()
    if kepada_custom:
        kepada_text = kepada_custom if kepada_custom.upper().startswith('YTH') else f"YTH. {kepada_custom}"
    terus_list = var_info.get("terus") or []

    # — Gunakan file template di folder dokum/ sesuai varian (bawaan maupun buatan admin) —
    dokum_dir = _dokum_dir()
    template_file = os.path.join(dokum_dir, var_info.get("filename") or "Lembar_Disposisi_Satsiber_TNI.docx")
    if not os.path.exists(template_file):
        template_file = os.path.join(dokum_dir, "Lembar_Disposisi_Satsiber_TNI.docx")
    doc = Document(template_file)

    def _fmt_tgl(v):
        if not v: return ""
        try:
            # terima format YYYY-MM-DD
            dt = datetime.datetime.strptime(str(v)[:10], "%Y-%m-%d")
            return dt.strftime("%d %b %Y")
        except: return str(v)

    def _clear_para(p):
        # hapus semua run/element di dalam <w:p> kecuali pPr
        pPr = p._p.find(qn('w:pPr'))
        for _ in range(len(list(p._p))):
            el = p._p[0]
            if el is not pPr:
                p._p.remove(el)
            else:
                break
        # pastikan pPr tetap (jika ada), sisanya sudah bersih
        # juga hapus jika ada sisa selain pPr
        for el in list(p._p):
            if el is not pPr:
                p._p.remove(el)

    def _add_run2(p, text, bold=False, sz_half=None, color=None):
        r = p.add_run(text)
        rPr = r._element.find(qn('w:rPr'))
        if rPr is None:
            rPr = r._element.makeelement(qn('w:rPr'), {})
            r._element.insert(0, rPr)
        rf = rPr.find(qn('w:rFonts'))
        if rf is None:
            rf = rPr.makeelement(qn('w:rFonts'), {}); rPr.append(rf)
        rf.set(qn('w:ascii'), 'Arial'); rf.set(qn('w:hAnsi'), 'Arial')
        if bold:
            b = rPr.find(qn('w:b'))
            if b is None:
                b = rPr.makeelement(qn('w:b'), {}); rPr.append(b)
        if color:
            cl = rPr.find(qn('w:color'))
            if cl is None:
                cl = rPr.makeelement(qn('w:color'), {}); rPr.append(cl)
            cl.set(qn('w:val'), color)
        if sz_half is not None:
            sz = rPr.find(qn('w:sz'))
            if sz is None:
                sz = rPr.makeelement(qn('w:sz'), {}); rPr.append(sz)
            sz.set(qn('w:val'), str(sz_half))
            szCs = rPr.find(qn('w:szCs'))
            if szCs is None:
                szCs = rPr.makeelement(qn('w:szCs'), {}); rPr.append(szCs)
            szCs.set(qn('w:val'), str(sz_half))
        return r

    def _br(p):
        r = p.add_run()
        r.add_break()

    table = doc.tables[0]    # — Isi tabel template (pertahankan border/margin/font asli) —
    paraf_text = (data.get("paraf") or "").strip()
    paraf_img = (data.get("paraf_image") or "").strip()
    has_img = bool(paraf_img and paraf_img.startswith("data:image"))

    # Row0 Col0: PARAF ... TANGGAL SURAT — persis template: label + titik dua
    # BOLD semua, jumlah spasi disamakan agar titik dua sejajar; nilai tidak bold.
    c0 = table.rows[0].cells[0]
    p0 = c0.paragraphs[0]
    _clear_para(p0)
    def _lab(teks_label, n_spasi, titik=': '):
        _add_run2(p0, teks_label + ' ' * n_spasi, bold=True, sz_half=17)
        _add_run2(p0, titik, bold=True, sz_half=17)
    _lab('PARAF', 25)
    if paraf_text:
        _add_run2(p0, paraf_text, sz_half=17)
    if has_img:
        try:
            header, b64 = paraf_img.split(",",1)
            img_bytes = base64.b64decode(b64)
            bio = io.BytesIO(img_bytes)
            # sisip spasi lalu gambar sejajar di samping tulisan PARAF
            _add_run2(p0, "  ", sz_half=17)
            rimg = p0.add_run()
            rimg.add_picture(bio, width=Inches(1.05))
        except Exception as e:
            print(f"[DISPOSISI] gagal paraf img: {e}")
    _br(p0); _br(p0)
    _lab('DITERIMA DARI', 10)
    _add_run2(p0, data.get("diterima_dari") or "", sz_half=17)
    _br(p0); _br(p0)
    _lab('TANGGAL TERIMA', 5)
    _add_run2(p0, _fmt_tgl(data.get("tanggal_terima")), sz_half=17)
    _br(p0)
    _lab('NO. AGENDA', 15, ':')
    _add_run2(p0, data.get("no_agenda") or "", sz_half=17)
    _br(p0)
    _lab('SURAT DARI', 16)
    _add_run2(p0, data.get("surat_dari") or "", sz_half=17)
    _br(p0)
    _lab('NOMOR SURAT', 11)
    _add_run2(p0, data.get("nomor_surat") or "", sz_half=17)
    _br(p0)
    _lab('JENIS SURAT', 13)
    _jenis_warna = {'sangat segera': 'FF0000', 'segera': 'B8860B', 'rahasia': '000000',
                    'dinas': '1D4ED8', 'undangan': '047857', 'laporan': '7C3AED'}.get(
                        (data.get("jenis_surat") or "").strip().lower(), '000000')
    _jenis_tebal = (data.get("jenis_surat") or "").strip().lower() in ('sangat segera', 'segera', 'rahasia')
    _add_run2(p0, (data.get("jenis_surat") or "").upper(), bold=_jenis_tebal, sz_half=17, color=_jenis_warna)
    _br(p0)
    _lab('TANGGAL SURAT', 6, ':')
    _add_run2(p0, ' ', sz_half=17)
    _add_run2(p0, _fmt_tgl(data.get("tanggal_surat")), sz_half=17)

    # Row0 Col1 (span 2): KEPADA — jangan dihapus/diubah, biarkan persis file docx yang dipilih (varian)
    if kepada_custom:
        c1 = table.rows[0].cells[1]
        p1 = c1.paragraphs[0]
        _clear_para(p1)
        _add_run2(p1, "KEPADA :", sz_half=18)
        _br(p1); _br(p1)
        _add_run2(p1, kepada_text, bold=True, sz_half=18)
    # else: biarkan sesuai template (tidak diubah)

    # Row1: PERIHAL (span 3)
    cr = table.rows[1].cells[0]
    pr = cr.paragraphs[0]
    _clear_para(pr)
    _add_run2(pr, "PERIHAL :  ", bold=True, sz_half=18)
    _add_run2(pr, data.get("perihal") or "", bold=True, sz_half=18)

    # Row4: checkbox lists (row index 4 in 0-based)
    # varian already menentukan terus_list yang ada di template, tapi kita isi ulang sesuai data
    terus_set = set(data.get("diteruskan_kepada") or [])
    isi_set = set(data.get("isi_disposisi") or [])
    # kolom 0: DITERUSKAN
    cc0 = table.rows[4].cells[0]
    pc0 = cc0.paragraphs[0]
    _clear_para(pc0)
    # isi sesuai terus_list varian, agar jumlah baris pas dengan template
    for i, label in enumerate(terus_list):
        box = "☑  " if label in terus_set else "☐  "
        _add_run2(pc0, box, sz_half=17)
        _add_run2(pc0, label, sz_half=16)
        if i != len(terus_list)-1:
            _br(pc0)
    # kolom 1 & 2: ISI DISPOSISI (template selalu 2 kolom isi)
    # Ambil daftar isi asli per varian = DISPOSISI_ISI_1/2 (sama semua)
    isi_all = DISPOSISI_ISI_1 + DISPOSISI_ISI_2
    # template membagi: kolom1 = isi1 (13), kolom2 = isi2 (13)
    for col_idx, lst in enumerate([DISPOSISI_ISI_1, DISPOSISI_ISI_2]):
        cc = table.rows[4].cells[1+col_idx]
        pc = cc.paragraphs[0]
        _clear_para(pc)
        for i, label in enumerate(lst):
            box = "☑  " if label in isi_set else "☐  "
            _add_run2(pc, box, sz_half=17)
            _add_run2(pc, label, sz_half=16)
            if i != len(lst)-1:
                _br(pc)

    # Row5: CATATAN (span 3)
    cr2 = table.rows[5].cells[0]
    pr2 = cr2.paragraphs[0]
    _clear_para(pr2)
    _add_run2(pr2, "CATATAN:", bold=True, sz_half=17)
    catatan = (data.get("catatan") or "").strip()
    if catatan:
        _add_run2(pr2, "  ", sz_half=17)
        _add_run2(pr2, catatan, sz_half=17)
    else:
        for _ in range(5):
            _br(pr2)

    return doc

@app.route('/api/disposisi/atasan-list', methods=['GET'])
@login_required
@disposisi_required
def disposisi_atasan_list():
    try:
        docs = fdb.collection('users').where('role', '==', 'atasan').stream()
        hasil=[]
        for d in docs:
            u=d.to_dict()
            j=u.get('jabatan_atasan') or ''
            hasil.append({'id': d.id, 'fullname': u.get('fullname') or '', 'jabatan_atasan': j, 'label': f"{u.get('fullname')} — {j}" if j else u.get('fullname')})
        hasil.sort(key=lambda x: x['fullname'].lower())
        return jsonify(hasil)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

# ===================== KELOLA FORMAT SURAT (ADMIN) =====================
# Admin upload .docx ke folder dokum/ + metadata (kepada, diteruskan). Otomatis
# muncul di dropdown user & atasan via GET /api/format-surat.
@app.route('/api/format-surat', methods=['GET'])
@login_required
@disposisi_required
def list_format_surat():
    try:
        semua = get_all_variants()
        hasil = [{'key': k, 'label': v.get('label'), 'kepada': v.get('kepada'),
                  'terus': v.get('terus') or [], 'filename': v.get('filename'),
                  'custom': bool(v.get('custom')), 'bawaan': (k in DISPOSISI_VARIANTS_BASE)}
                 for k, v in semua.items()]
        hasil.sort(key=lambda x: (not x['bawaan'], x['label'] or ''))
        return jsonify(hasil)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/api/format-surat', methods=['POST'])
@admin_required
def tambah_format_surat():
    try:
        nama = (request.form.get('nama') or request.form.get('label') or '').strip()
        kepada = (request.form.get('kepada') or '').strip()
        terus_raw = request.form.get('terus') or request.form.get('diteruskan') or ''
        key_req = (request.form.get('key') or '').strip().lower()
        if 'file' not in request.files or not request.files['file'].filename:
            return jsonify({"error": "File template .docx wajib diupload."}), 400
        f = request.files['file']
        if '.' not in f.filename or f.filename.rsplit('.', 1)[1].lower() != 'docx':
            return jsonify({"error": "File harus berformat .docx."}), 400
        if not nama:
            nama = os.path.splitext(f.filename)[0]
        key = key_req or _slugify_key(nama)
        if key in DISPOSISI_VARIANTS_BASE:
            return jsonify({"error": f"Kode '{key}' sudah dipakai format bawaan."}), 400
        # parse daftar diteruskan: pisahkan koma / baris baru / JSON array
        terus = []
        if terus_raw.strip().startswith('['):
            import json as _json
            try:
                terus = _json.loads(terus_raw)
            except Exception:
                terus = []
        else:
            import re as _re
            terus = [t.strip().upper() for t in _re.split(r'[\n,;]+', terus_raw) if t.strip()]
        if not terus:
            terus = list(DISPOSISI_VARIANTS_BASE['dansat']['terus'])
        filename = secure_filename(f.filename)
        if not filename.lower().endswith('.docx'):
            filename += '.docx'
        # hindari timpa file bawaan
        dokum = _dokum_dir()
        os.makedirs(dokum, exist_ok=True)
        if os.path.exists(os.path.join(dokum, filename)):
            filename = f"{key}_{filename}"
        f.save(os.path.join(dokum, filename))
        if not kepada:
            kepada = f"YTH. {nama.upper()}"
        elif not kepada.upper().startswith('YTH'):
            kepada = f"YTH. {kepada}"
        label = f"{nama} ({filename})"
        fdb.collection('format_surat').document(key).set({
            'key': key, 'label': label, 'nama': nama, 'kepada': kepada,
            'terus': terus, 'filename': filename,
            'dibuat_oleh': session.get('user_name', 'Admin'),
            'dibuat_pada': datetime.datetime.now().strftime('%d %b %Y %H:%M'),
        })
        return jsonify({"message": f"Format '{nama}' tersimpan di dokum/{filename} dan tampil di opsi user & atasan.", "key": key}), 201
    except Exception as e:
        import traceback; traceback.print_exc()
        return jsonify({"error": f"Gagal menyimpan format: {str(e)}"}), 500

@app.route('/api/format-surat/<string:key>', methods=['DELETE'])
@admin_required
def hapus_format_surat(key):
    key = (key or '').strip().lower()
    if key in DISPOSISI_VARIANTS_BASE:
        return jsonify({"error": "Format bawaan tidak boleh dihapus."}), 400
    try:
        doc_ref = fdb.collection('format_surat').document(key)
        doc = doc_ref.get()
        filename = (doc.to_dict() or {}).get('filename') if doc.exists else None
        if not filename:
            # mungkin file auto-scan tanpa metadata: cari key cocok? tetap tolak aman
            return jsonify({"error": "Format tidak ditemukan."}), 404
        fpath = os.path.join(_dokum_dir(), filename)
        if os.path.exists(fpath):
            os.remove(fpath)
        doc_ref.delete()
        return jsonify({"message": f"Format '{key}' dihapus (file dokum/{filename})."}), 200
    except Exception as e:
        return jsonify({"error": f"Gagal menghapus: {str(e)}"}), 500

@app.route('/admin/format-surat.html')
@admin_required
def serve_admin_format_surat():
    return render_template('admin/format-surat.html', user_name=session.get('user_name', 'Admin'))

@app.route('/api/disposisi/generate', methods=['POST'])
@login_required
@disposisi_required
def generate_disposisi():
    try:
        if _unduh_ditolak():
            return jsonify({"error": "Akun ini hanya dapat melihat Jejak dan Rantai."}), 403
        data = _normalisasi_data_disposisi()
        doc = _build_disposisi_docx(data)
        import io
        buf = io.BytesIO(); doc.save(buf); buf.seek(0)
        fname = f"Lembar_Disposisi_{(data.get('no_agenda') or data.get('nomor_surat') or 'Satsiber').replace('/','-')}.docx"
        return send_file(buf, as_attachment=True, download_name=fname, mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
    except Exception as e:
        import traceback; traceback.print_exc()
        return jsonify({"error": f"Gagal generate disposisi: {str(e)}"}), 500


def _normalisasi_data_disposisi():
    """Samakan bentuk data disposisi dari JSON maupun form (dipakai generate Word & PDF)."""
    data = request.get_json(silent=True) or request.form.to_dict()
    # normalise lists from form (checkbox arrays)
    if isinstance(data.get('diteruskan_kepada'), str):
        import json as _json
        try:
            v = _json.loads(data['diteruskan_kepada'])
            data['diteruskan_kepada'] = v if isinstance(v, list) else [data['diteruskan_kepada']]
        except: data['diteruskan_kepada'] = [data['diteruskan_kepada']] if data['diteruskan_kepada'] else []
    if isinstance(data.get('isi_disposisi'), str):
        import json as _json
        try:
            v = _json.loads(data['isi_disposisi'])
            data['isi_disposisi'] = v if isinstance(v, list) else [data['isi_disposisi']]
        except: data['isi_disposisi'] = [data['isi_disposisi']] if data['isi_disposisi'] else []
    # support separate form fields: diteruskan_kepada[] etc.
    for key in list(request.form.keys()):
        if key.startswith('diteruskan_kepada'):
            data['diteruskan_kepada'] = request.form.getlist(key)
        if key.startswith('isi_disposisi'):
            data['isi_disposisi'] = request.form.getlist(key)

    # Fallback: getlist for json? try to read lists from request.values
    if not data.get('diteruskan_kepada'):
        lst = request.form.getlist('diteruskan_kepada')
        if lst: data['diteruskan_kepada'] = lst
    if not data.get('isi_disposisi'):
        lst = request.form.getlist('isi_disposisi')
        if lst: data['isi_disposisi'] = lst
    return data


@app.route('/api/disposisi/generate-pdf', methods=['POST'])
@login_required
@disposisi_required
def generate_disposisi_pdf():
    """Cetak Lembar Disposisi sebagai PDF via server (MS Word).

    Hasil dijamin sama persis dengan versi Word karena dikonversi langsung
    dari dokumen yang sama — tidak bergantung pada screenshot browser
    (html2canvas) yang sering gagal.
    """
    import io, os, tempfile, uuid
    docx_path, pdf_path = None, None
    try:
        if _unduh_ditolak():
            return jsonify({"error": "Akun ini hanya dapat melihat Jejak dan Rantai."}), 403
        data = _normalisasi_data_disposisi()
        doc = _build_disposisi_docx(data)
        tmpdir = tempfile.mkdtemp(prefix='disposisi_')
        tag = uuid.uuid4().hex[:8]
        docx_path = os.path.join(tmpdir, f'disposisi_{tag}.docx')
        pdf_path = os.path.join(tmpdir, f'disposisi_{tag}.pdf')
        doc.save(docx_path)
        # CoInitialize + lock: Flask jalan di thread, Word COM tidak tahan paralel
        try:
            import pythoncom as _pc
            _pc.CoInitialize()
        except Exception:
            pass
        try:
            with _WORD2PDF_LOCK:
                from docx2pdf import convert as _docx2pdf
                _docx2pdf(docx_path, pdf_path)
        finally:
            try:
                import pythoncom as _pc2
                _pc2.CoUninitialize()
            except Exception:
                pass
        if not os.path.exists(pdf_path) or os.path.getsize(pdf_path) < 100:
            return jsonify({"error": "Konversi gagal menghasilkan PDF."}), 500
        with open(pdf_path, 'rb') as fh:
            buf = io.BytesIO(fh.read())
        buf.seek(0)
        fname = f"Lembar_Disposisi_{(data.get('no_agenda') or data.get('nomor_surat') or 'Satsiber').replace('/','-')}.pdf"
        return send_file(buf, as_attachment=True, download_name=fname, mimetype='application/pdf')
    except Exception as e:
        import traceback; traceback.print_exc()
        return jsonify({"error": f"Gagal membuat PDF disposisi: {str(e)}"}), 500
    finally:
        for _p in (docx_path, pdf_path):
            try:
                if _p and os.path.exists(_p):
                    os.remove(_p)
            except OSError:
                pass
        try:
            if docx_path and os.path.isdir(os.path.dirname(docx_path)):
                os.rmdir(os.path.dirname(docx_path))
        except OSError:
            pass

# ===================== RIWAYAT INPUT DISPOSISI (bersama semua user, tanpa batas) =====================
# Setiap kolom form punya satu dokumen (id = id kolom) berisi array `items`
# (nilai terbaru di depan). Tidak ada batas jumlah — semua user membaca &
# menulis ke tempat yang sama sehingga riwayat bersifat kolektif.
RIWAYAT_DISPOSISI_FIELDS = {'d-paraf', 'd-diterima-dari', 'd-no-agenda', 'd-surat-dari',
                            'd-nomor-surat', 'd-perihal', 'd-catatan'}

@app.route('/api/disposisi/history', methods=['GET'])
@login_required
@disposisi_required
def get_disposisi_history():
    try:
        hasil = {}
        for d in fdb.collection('disposisi_hist').stream():
            items = (d.to_dict() or {}).get('items') or []
            if items:
                hasil[d.id] = items
        return jsonify(hasil)
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route('/api/disposisi/history', methods=['POST'])
@login_required
@disposisi_required
def save_disposisi_history():
    data = request.get_json(silent=True) or {}
    items = data.get('items') or {}
    bersih = {}
    for fid, vals in items.items():
        if fid not in RIWAYAT_DISPOSISI_FIELDS or not isinstance(vals, list):
            continue
        unik = []
        for v in vals:
            v = str(v or '').strip()
            if v and v not in unik:
                unik.append(v)
        if unik:
            bersih[fid] = unik
    if not bersih:
        return jsonify({"message": "Tidak ada riwayat baru."}), 200
    try:
        @fdb_transactional
        def _merge(tx):
            # WAJIB: semua pembacaan dulu sebelum ada penulisan (pola transaksi)
            snaps = {}
            for fid in bersih:
                ref = fdb.collection('disposisi_hist').document(fid)
                snaps[fid] = (ref, ref.get(transaction=tx))
            for fid, vals in bersih.items():
                ref, snap = snaps[fid]
                lama = ((snap.to_dict() or {}).get('items') or []) if snap.exists else []
                gabung = list(vals) + [x for x in lama if x not in vals]
                tx.set(ref, {'items': gabung,
                             'diperbarui_pada': SERVER_TIMESTAMP,
                             'diperbarui_oleh': session.get('user_name')})
        _merge(fdb.transaction())
        return jsonify({"message": "Riwayat tersimpan."}), 200
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route('/api/disposisi/history/<string:field_id>', methods=['DELETE'])
@admin_required
def clear_disposisi_history(field_id):
    if field_id not in RIWAYAT_DISPOSISI_FIELDS:
        return jsonify({"error": "Kolom tidak dikenal."}), 400
    try:
        fdb.collection('disposisi_hist').document(field_id).delete()
        return jsonify({"message": "Riwayat kolom dihapus untuk semua pengguna."}), 200
    except Exception as e:
        return jsonify({"error": str(e)}), 500
# =====================================================================================================

# ===================== KIRIM SURAT DISPOSISI KE ATASAN (SURAT MASUK) =====================
DISPOSISI_UPLOAD_FOLDER = os.path.join(UPLOAD_FOLDER, 'disposisi')
os.makedirs(DISPOSISI_UPLOAD_FOLDER, exist_ok=True)

@app.route('/api/disposisi/kirim', methods=['POST'])
@login_required
@disposisi_required
def kirim_disposisi():
    try:
        # Pantau-only (atasan tingkat 1-2 & 4+) tidak boleh mengirim
        if not _boleh_mengirim():
            return jsonify({"error": "Akun ini hanya dapat memantau. Pengiriman dilakukan oleh persuratan."}), 403
        # Terima JSON maupun multipart/form-data (untuk lampiran PDF surat undangan)
        if request.content_type and 'multipart/form-data' in request.content_type:
            data = request.form.to_dict()
            import json as _json
            for _k in ('diteruskan_kepada', 'isi_disposisi'):
                if isinstance(data.get(_k), str):
                    try:
                        _v = _json.loads(data[_k])
                        data[_k] = _v if isinstance(_v, list) else [data[_k]]
                    except:
                        pass
            for _k in ('diteruskan_kepada', 'isi_disposisi'):
                _lst = request.form.getlist(_k)
                if len(_lst) > 1:
                    data[_k] = _lst
        else:
            data = request.get_json(silent=True) or {}
        # Tujuan bisa BANYAK atasan: `atasan_ids` (list/JSON) atau tunggal `atasan_id` (kompatibel lama)
        import json as _json2
        atasan_ids = []
        _mentah = data.get('atasan_ids')
        if isinstance(_mentah, list):
            atasan_ids = [str(x or '').strip() for x in _mentah]
        elif isinstance(_mentah, str) and _mentah.strip():
            try:
                _v = _json2.loads(_mentah)
                atasan_ids = [str(x or '').strip() for x in (_v if isinstance(_v, list) else [_mentah])]
            except:
                atasan_ids = [s.strip() for s in _mentah.split(',')]
        if not atasan_ids and isinstance(request.form.getlist('atasan_ids'), list):
            _lst = [s.strip() for s in request.form.getlist('atasan_ids') if (s or '').strip()]
            if len(_lst) > 1 or (_lst and _lst[0].startswith('[')):
                try:
                    _v = _json2.loads(_lst[0]) if len(_lst) == 1 else _lst
                    atasan_ids = [str(x or '').strip() for x in (_v if isinstance(_v, list) else _lst)]
                except:
                    atasan_ids = _lst
        if not atasan_ids:
            tunggal = (data.get('atasan_id') or data.get('atasan_select') or '').strip()
            if tunggal:
                atasan_ids = [tunggal]
        # buang kosong & duplikat (pertahankan urutan pilih)
        atasan_ids = list(dict.fromkeys([a for a in atasan_ids if a]))
        if not atasan_ids:
            return jsonify({"error": "Pilih minimal satu atasan tujuan pada 'Ditujukan Kepada'."}), 400
        # validasi semua atasan (ada & berperar Atasan)
        daftar_atasan = []
        for _aid in atasan_ids:
            _doc = fdb.collection('users').document(_aid).get()
            if not _doc.exists:
                return jsonify({"error": "Salah satu atasan tujuan tidak ditemukan."}), 404
            _a = _doc.to_dict()
            if _a.get('role') != 'atasan':
                return jsonify({"error": f"{_a.get('fullname') or 'Tujuan'} bukan akun dengan peran Atasan."}), 400
            daftar_atasan.append((_aid, _a))

        pengirim_id = session.get('user_id')
        pengirim_nama = session.get('user_name') or (fdb.collection('users').document(pengirim_id).get().to_dict() or {}).get('fullname','')
        # KEPADA ikut format docx template (varian), bukan nama personal — kosongkan override
        data['kepada_custom'] = ""
        # build docx sesuai template docx yang dipilih (margin/border/font 100% identik)
        doc = _build_disposisi_docx(data)
        fname_base = f"Lembar_Disposisi_{(data.get('no_agenda') or data.get('nomor_surat') or 'Satsiber').replace('/','-')}_{datetime.datetime.now().strftime('%Y%m%d%H%M%S')}"
        fname = fname_base + ".docx"
        fpath = os.path.join(DISPOSISI_UPLOAD_FOLDER, fname)
        doc.save(fpath)

        # — Lampiran PDF surat undangan dari pengirim pertama (user/staf) —
        lampiran_simpan, lampiran_asli = '', ''
        lampiran_file = request.files.get('lampiran_pdf') or request.files.get('lampiran') or request.files.get('file')
        if lampiran_file and lampiran_file.filename:
            nama_asli = lampiran_file.filename
            if '.' not in nama_asli or nama_asli.rsplit('.', 1)[1].lower() != 'pdf':
                return jsonify({"error": "Lampiran surat undangan harus berformat PDF."}), 400
            lampiran_file.seek(0, os.SEEK_END)
            ukuran_mb = lampiran_file.tell() / (1024 * 1024)
            lampiran_file.seek(0)
            if ukuran_mb > 10:
                return jsonify({"error": "Ukuran PDF lampiran maksimal 10 MB."}), 400
            lampiran_simpan = f"lampiran_{datetime.datetime.now().strftime('%Y%m%d%H%M%S')}_{secure_filename(nama_asli)}"
            lampiran_file.save(os.path.join(DISPOSISI_UPLOAD_FOLDER, lampiran_simpan))
            lampiran_asli = nama_asli

        # simpan ke database: satu baris per atasan tujuan (berkas sama dipakai bersama)
        _public = get_cloudflared_url()
        _link = f'{_public}/USER/alat-bantuan.html' if _public else ''
        _lamp_path = os.path.join(DISPOSISI_UPLOAD_FOLDER, lampiran_simpan) if lampiran_simpan else ''
        ids_terkirim, nama_terkirim = [], []
        tg_terkirim = 0
        for atasan_id, atasan in daftar_atasan:
            ref = fdb.collection('disposisi_surat').document()
            ref.set({
            'pengirim_id': pengirim_id,
            'pengirim_nama': pengirim_nama,
            'atasan_id': atasan_id,
            'atasan_nama': atasan.get('fullname') or '',
            'atasan_jabatan': atasan.get('jabatan_atasan') or '',
            'varian': data.get('varian') or 'dansat',
            'kepada_custom': "",  # ikut format docx template varian
            'perihal': data.get('perihal') or '',
            'nomor_surat': data.get('nomor_surat') or '',
            'no_agenda': data.get('no_agenda') or '',
            'surat_dari': data.get('surat_dari') or '',
            'jenis_surat': data.get('jenis_surat') or '',
            'tanggal_surat': data.get('tanggal_surat') or '',
            'tanggal_terima': data.get('tanggal_terima') or '',
            'diteruskan_kepada': data.get('diteruskan_kepada') or [],
            'isi_disposisi': data.get('isi_disposisi') or [],
            'catatan': data.get('catatan') or '',
            'paraf': data.get('paraf') or '',
            'paraf_image': data.get('paraf_image') or '',
            'file_nama_simpan': fname,
            'file_nama_asli': fname,
            'lampiran_nama_simpan': lampiran_simpan,
            'lampiran_nama_asli': lampiran_asli,
            'created_at': datetime.datetime.now().strftime('%d %b %Y %H:%M'),
            'created_ts': SERVER_TIMESTAMP,
            'status': 'baru',
            'dibaca_oleh': []
            })
            ids_terkirim.append(ref.id)
            nama_terkirim.append(atasan.get('fullname') or '')
            try:
                if notifikasi_disposisi_telegram(atasan, {
                        'perihal': data.get('perihal') or '',
                        'nomor_surat': data.get('nomor_surat') or '',
                        'no_agenda': data.get('no_agenda') or '',
                        'jenis_surat': data.get('jenis_surat') or '',
                        'pengirim_nama': pengirim_nama,
                        'file_docx': fpath,
                        'lampiran_pdf': _lamp_path,
                        'app_link': _link}):
                    tg_terkirim += 1
            except Exception as _e:
                print(f'[TELEGRAM] notifikasi disposisi gagal: {_e}')
        if len(ids_terkirim) == 1:
            _msg = f"Surat disposisi berhasil dikirim ke {nama_terkirim[0]}."
            if tg_terkirim:
                _msg += " Notifikasi Telegram terkirim."
            return jsonify({"message": _msg, "id": ids_terkirim[0], "telegram_terkirim": tg_terkirim}), 201
        _msg = f"Surat disposisi berhasil dikirim ke {len(ids_terkirim)} atasan: {', '.join(nama_terkirim)}."
        if tg_terkirim:
            _msg += f" Notifikasi Telegram terkirim ke {tg_terkirim} penerima."
        return jsonify({"message": _msg, "ids": ids_terkirim, "telegram_terkirim": tg_terkirim}), 201
    except Exception as e:
        import traceback; traceback.print_exc()
        return jsonify({"error": f"Gagal mengirim disposisi: {str(e)}"}), 500

@app.route('/api/disposisi/surat-masuk', methods=['GET'])
@atasan_required
def list_surat_masuk():
    try:
        uid = session.get('user_id')
        role = session.get('user_role')
        if role == 'admin':
            docs = list(fdb.collection('disposisi_surat').stream())
        else:
            docs = list(fdb.collection('disposisi_surat').where('atasan_id', '==', uid).stream())
        return jsonify(_serialize_surat_masuk(docs))
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/api/disposisi/surat-untuk-saya', methods=['GET'])
@login_required
@disposisi_required
def surat_untuk_saya():
    """Kotak masuk disposisi untuk SEMUA peran (termasuk user/staf biasa).

    Atasan memakai /api/disposisi/surat-masuk; endpoint ini proporsional
    untuk penerima non-atasan hasil terusan (polling popup + unduh).
    """
    try:
        uid = session.get('user_id')
        if session.get('user_role') == 'admin':
            docs = list(fdb.collection('disposisi_surat').stream())
        else:
            docs = list(fdb.collection('disposisi_surat').where('atasan_id', '==', uid).stream())
        return jsonify(_serialize_surat_masuk(docs))
    except Exception as e:
        return jsonify({"error": str(e)}), 500

def _serialize_surat_masuk(docs):
    saya = session.get('user_id')
    hasil = []
    for d in docs:
        x = d.to_dict()
        hasil.append({
            'id': d.id,
            'pengirim_nama': x.get('pengirim_nama'),
            'pengirim_id': x.get('pengirim_id'),
            'atasan_nama': x.get('atasan_nama'),
            'perihal': x.get('perihal'),
            'nomor_surat': x.get('nomor_surat'),
            'no_agenda': x.get('no_agenda'),
            'surat_dari': x.get('surat_dari'),
            'jenis_surat': x.get('jenis_surat') or '',
            'marker': marker_jenis(x.get('jenis_surat')),
            'created_at': x.get('created_at'),
            'file_nama_asli': x.get('file_nama_asli'),
            'lampiran_nama_asli': x.get('lampiran_nama_asli') or '',
                'ada_lampiran': bool(x.get('lampiran_nama_simpan')),
                'status': x.get('status'),
                'sudah_dibaca': bool(saya) and saya in (x.get('dibaca_oleh') or []),
                'varian': x.get('varian'),
                'konfirmasi_oleh_nama': x.get('konfirmasi_oleh_nama') or '',
                'konfirmasi_pada': x.get('konfirmasi_pada') or ''
            })
    # terbaru dulu
    hasil.sort(key=lambda h: h['created_at'] or '', reverse=True)
    return hasil

@app.route('/api/disposisi/surat-masuk/<string:surat_id>/baca', methods=['POST'])
@login_required
@disposisi_required
def tandai_surat_dibaca(surat_id):
    """Tandai surat sudah dibaca oleh saya (agar popup tidak muncul lagi)."""
    try:
        doc_ref = fdb.collection('disposisi_surat').document(surat_id)
        doc = doc_ref.get()
        if not doc.exists:
            return jsonify({"error": "Surat tidak ditemukan."}), 404
        x = doc.to_dict()
        saya = session.get('user_id')
        boleh = (x.get('atasan_id') == saya or x.get('pengirim_id') == saya
                 or session.get('user_role') == 'admin')
        if not boleh:
            return jsonify({"error": "Anda tidak memiliki akses."}), 403
        doc_ref.update({'dibaca_oleh': ArrayUnion([saya])})
        return jsonify({"message": "Ditandai dibaca."}), 200
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/api/disposisi/surat-masuk/<string:surat_id>/download', methods=['GET'])
@login_required
@disposisi_required
def download_surat_masuk(surat_id):
    # Mode LIHAT (?lihat=1): atasan pantau-only (tingkat 1-2) boleh MEMBUKA
    # isi/file di browser (inline). Tanpa itu mereka tetap ditolak (403)
    # agar tidak bisa mengunduh/menyebarkan berkas.
    mode_lihat = (request.args.get('lihat') == '1')
    if _unduh_ditolak(mode_lihat):
        return jsonify({"error": "Akun ini hanya dapat melihat Jejak dan Rantai."}), 403
    doc = fdb.collection('disposisi_surat').document(surat_id).get()
    if not doc.exists:
        return jsonify({"error": "Surat tidak ditemukan."}), 404
    x=doc.to_dict()
    saya = session.get('user_id')
    peran = session.get('user_role')
    boleh = (x.get('atasan_id') == saya or x.get('pengirim_id') == saya or peran == 'admin')
    if not boleh:
        return jsonify({"error": "Anda tidak memiliki akses ke surat ini."}), 403
    fname = x.get('file_nama_simpan')
    if not fname:
        return jsonify({"error": "File tidak tersedia."}), 404
    fpath = os.path.join(DISPOSISI_UPLOAD_FOLDER, fname)
    if not os.path.exists(fpath):
        return jsonify({"error": "File fisik tidak ditemukan di server."}), 404
    _catat_log_disposisi(surat_id, 'lihat' if mode_lihat else 'unduh')
    return send_from_directory(DISPOSISI_UPLOAD_FOLDER, fname, as_attachment=not mode_lihat, download_name=x.get('file_nama_asli') or fname)

@app.route('/api/disposisi/surat-masuk/<string:surat_id>', methods=['GET'])
@login_required
@disposisi_required
def detail_surat_masuk(surat_id):
    doc = fdb.collection('disposisi_surat').document(surat_id).get()
    if not doc.exists:
        return jsonify({"error": "Surat tidak ditemukan."}), 404
    x=doc.to_dict()
    saya = session.get('user_id')
    peran = session.get('user_role')
    boleh = (x.get('atasan_id') == saya or x.get('pengirim_id') == saya or peran == 'admin')
    if not boleh:
        return jsonify({"error": "Anda tidak memiliki akses."}), 403
    # kembalikan semua field untuk prefill teruskan
    return jsonify({
        'id': doc.id,
        'pengirim_nama': x.get('pengirim_nama'),
        'perihal': x.get('perihal'),
        'nomor_surat': x.get('nomor_surat'),
        'no_agenda': x.get('no_agenda'),
        'surat_dari': x.get('surat_dari'),
        'jenis_surat': x.get('jenis_surat'),
        'tanggal_surat': x.get('tanggal_surat'),
        'tanggal_terima': x.get('tanggal_terima'),
        'diterima_dari': x.get('diterima_dari'),
        'diteruskan_kepada': x.get('diteruskan_kepada'),
        'isi_disposisi': x.get('isi_disposisi'),
        'catatan': x.get('catatan'),
        'paraf': x.get('paraf'),
        'paraf_image': x.get('paraf_image') or '',
        'varian': x.get('varian'),
        'kepada_custom': x.get('kepada_custom'),
        'marker': marker_jenis(x.get('jenis_surat')),
        'lampiran_nama_asli': x.get('lampiran_nama_asli') or '',
        'ada_lampiran': bool(x.get('lampiran_nama_simpan')),
        'konfirmasi_oleh_nama': x.get('konfirmasi_oleh_nama') or '',
        'konfirmasi_pada': x.get('konfirmasi_pada') or '',
        'created_at': x.get('created_at')
    })

@app.route('/api/disposisi/surat-masuk/<string:surat_id>/lampiran', methods=['GET'])
@login_required
@disposisi_required
def download_lampiran_disposisi(surat_id):
    # Sama seperti download: ?lihat=1 membuka PDF inline untuk pantau-only.
    mode_lihat = (request.args.get('lihat') == '1')
    if _unduh_ditolak(mode_lihat):
        return jsonify({"error": "Akun ini hanya dapat melihat Jejak dan Rantai."}), 403
    doc = fdb.collection('disposisi_surat').document(surat_id).get()
    if not doc.exists:
        return jsonify({"error": "Surat tidak ditemukan."}), 404
    x = doc.to_dict()
    saya = session.get('user_id')
    peran = session.get('user_role')
    boleh = (x.get('atasan_id') == saya or x.get('pengirim_id') == saya or peran == 'admin')
    if not boleh:
        return jsonify({"error": "Anda tidak memiliki akses ke lampiran ini."}), 403
    fname = x.get('lampiran_nama_simpan')
    if not fname:
        return jsonify({"error": "Tidak ada lampiran PDF pada surat ini."}), 404
    fpath = os.path.join(DISPOSISI_UPLOAD_FOLDER, fname)
    if not os.path.exists(fpath):
        return jsonify({"error": "File lampiran tidak ditemukan di server."}), 404
    _catat_log_disposisi(surat_id, 'lihat-lampiran' if mode_lihat else 'lampiran')
    return send_from_directory(DISPOSISI_UPLOAD_FOLDER, fname, as_attachment=False,
                               download_name=x.get('lampiran_nama_asli') or fname,
                               mimetype='application/pdf')

@app.route('/api/disposisi/bawahan-list', methods=['GET'])
@login_required
@disposisi_required
def bawahan_list():
    try:
        uid = session.get('user_id')
        # bawahan langsung: users where superior_id == uid
        docs = list(fdb.collection('users').where('superior_id', '==', uid).stream())
        hasil=[]
        for d in docs:
            u=d.to_dict()
            hasil.append({'id': d.id, 'fullname': u.get('fullname') or '', 'role': u.get('role') or '', 'jabatan_atasan': u.get('jabatan_atasan') or '', 'bagian': u.get('bagian') or '', 'email': u.get('email') or ''})
        hasil.sort(key=lambda x: x['fullname'].lower())
        return jsonify(hasil)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/api/disposisi/staf-list', methods=['GET'])
@login_required
@disposisi_required
def staf_list():
    """Daftar akun staf/persuratan (role=user) sebagai tujuan teruskan.
    Dibutuhkan karena banyak akun user belum punya superior_id sehingga
    tidak muncul di bawahan-list."""
    try:
        docs = list(fdb.collection('users').where('role', '==', 'user').stream())
        hasil=[]
        for d in docs:
            u=d.to_dict()
            hasil.append({'id': d.id, 'fullname': u.get('fullname') or '', 'role': u.get('role') or '', 'jabatan_atasan': u.get('jabatan_atasan') or '', 'bagian': u.get('bagian') or '', 'email': u.get('email') or ''})
        hasil.sort(key=lambda x: x['fullname'].lower())
        return jsonify(hasil)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/api/disposisi/teruskan/<string:surat_id>', methods=['POST'])
@login_required
@disposisi_required
def teruskan_disposisi(surat_id):
    try:
        # Atasan (semua tingkat), staf persuratan, dan admin berhak MENERUSKAN
        # surat yang dipegangnya. Aturan pantau-only hanya membatasi pembuatan
        # surat baru (lihat /api/disposisi/kirim), bukan penerusan disposisi.
        saya = session.get('user_id')
        if session.get('user_role') not in ('atasan', 'admin') and not is_persuratan(saya):
            return jsonify({"error": "Akun ini tidak berhak meneruskan disposisi."}), 403
        orig_doc = fdb.collection('disposisi_surat').document(surat_id).get()
        if not orig_doc.exists:
            return jsonify({"error": "Surat asal tidak ditemukan."}), 404
        orig = orig_doc.to_dict()
        # hanya PEMEGANG surat (atasan_id == saya) atau admin yang boleh
        # meneruskan — mencakup relay user persuratan dalam rantai berjenjang;
        # khusus persuratan: boleh meneruskan surat yang sedang dipegang atasan
        # tingkat 1-2 yang dipegang (mereka hanya memantau, tangannya persuratan);
        # pengirim surat boleh meneruskan ulang dari tab Terkirim.
        _pemegang = orig.get('atasan_id')
        _boleh = (_pemegang == saya or orig.get('pengirim_id') == saya or session.get('user_role') == 'admin')
        if not _boleh and is_persuratan(saya):
            _t_pemegang = tingkat_of(_pemegang)
            if _t_pemegang is not None and _t_pemegang <= 2:
                _boleh = True
        if not _boleh:
            return jsonify({"error": "Anda bukan penerima/pemegang surat ini."}), 403
        # Terima JSON maupun multipart/form-data (untuk upload PDF lampiran)
        if request.content_type and 'multipart/form-data' in request.content_type:
            data = request.form.to_dict()
            import json as _json
            for _k in ('diteruskan_kepada', 'isi_disposisi'):
                if isinstance(data.get(_k), str):
                    try:
                        _v = _json.loads(data[_k])
                        data[_k] = _v if isinstance(_v, list) else [data[_k]]
                    except:
                        pass
            for _k in ('diteruskan_kepada', 'isi_disposisi'):
                _lst = request.form.getlist(_k)
                # getlist mengembalikan semua nilai; abaikan jika hanya 1 string JSON
                if len(_lst) > 1 or (len(_lst) == 1 and not isinstance(data.get(_k), list)):
                    # jika form mengirim JSON-string tunggal, _lst berupa [json]; pakai data yg sudah diparsing
                    if len(_lst) == 1 and isinstance(data.get(_k), list):
                        pass
                    elif not (len(_lst) == 1 and _lst[0].startswith('[')):
                        data[_k] = _lst
        else:
            data = request.get_json(silent=True) or {}
        # Tujuan bisa BANYAK: `tujuan_ids` (list/JSON) atau tunggal `tujuan_id` (kompatibel lama)
        import json as _json3
        tujuan_ids = []
        _mentah3 = data.get('tujuan_ids')
        if isinstance(_mentah3, list):
            tujuan_ids = [str(x or '').strip() for x in _mentah3]
        elif isinstance(_mentah3, str) and _mentah3.strip():
            try:
                _v3 = _json3.loads(_mentah3)
                tujuan_ids = [str(x or '').strip() for x in (_v3 if isinstance(_v3, list) else [_mentah3])]
            except:
                tujuan_ids = [s.strip() for s in _mentah3.split(',')]
        if not tujuan_ids:
            _lst3 = [s.strip() for s in request.form.getlist('tujuan_ids') if (s or '').strip()]
            if _lst3:
                try:
                    _v3 = _json3.loads(_lst3[0]) if len(_lst3) == 1 else _lst3
                    tujuan_ids = [str(x or '').strip() for x in (_v3 if isinstance(_v3, list) else _lst3)]
                except:
                    tujuan_ids = _lst3
        if not tujuan_ids:
            _tunggal = (data.get('tujuan_id') or data.get('atasan_id') or '').strip()
            if _tunggal:
                tujuan_ids = [_tunggal]
        tujuan_ids = list(dict.fromkeys([t for t in tujuan_ids if t]))
        if not tujuan_ids:
            # Atasan tingkat 1-2 tanpa pilihan tujuan -> otomatis ke SEMUA
            # akun persuratan (langsung ke bagian persuratan, tanpa opsi).
            if session.get('user_role') == 'atasan':
                _tkt0 = tingkat_of(saya)
                if _tkt0 is not None and _tkt0 <= 2:
                    try:
                        for _d in fdb.collection('users').where('role', '==', 'user').stream():
                            _u = _d.to_dict() or {}
                            if _u.get('bagian') == 'persuratan':
                                tujuan_ids.append(_d.id)
                        tujuan_ids = list(dict.fromkeys(tujuan_ids))
                    except Exception:
                        pass
            if not tujuan_ids:
                return jsonify({"error": "Pilih minimal satu tujuan teruskan."}), 400
        daftar_tujuan = []
        for _tid in tujuan_ids:
            _tdoc = fdb.collection('users').document(_tid).get()
            if not _tdoc.exists:
                return jsonify({"error": "Salah satu tujuan tidak ditemukan."}), 404
            daftar_tujuan.append((_tid, _tdoc.to_dict()))
        # validasi bawahan langsung (opsional: longgarkan untuk admin)
        for _tid, _t in daftar_tujuan:
            if _t.get('superior_id') != saya and session.get('user_role') != 'admin':
                # tetap izinkan jika role bawahan adalah user/atasan tapi superior bukan saya? Beri warning tapi izinkan untuk fleksibilitas
                pass
        # CATATAN: tidak ada larangan tujuan (sesuai permintaan pemilik):
        # atasan tingkat berapa pun bebas meneruskan ke siapa saja;
        # UI mengarahkan atasan 1-2 langsung ke bagian persuratan.
        # data baru: gabung orig + override dari form teruskan
        varian, _, _ = get_variant_info(data.get('varian') or orig.get('varian') or 'dansat')
        # override fields jika ada
        fields = ['paraf','paraf_image','diterima_dari','tanggal_terima','no_agenda','surat_dari','nomor_surat','jenis_surat','tanggal_surat','perihal','catatan']
        new_data = {}
        for f in fields:
            new_data[f] = data.get(f) if data.get(f) is not None else orig.get(f)
        new_data['varian'] = varian
        new_data['diteruskan_kepada'] = data.get('diteruskan_kepada') if isinstance(data.get('diteruskan_kepada'), list) else orig.get('diteruskan_kepada') or []
        new_data['isi_disposisi'] = data.get('isi_disposisi') if isinstance(data.get('isi_disposisi'), list) else orig.get('isi_disposisi') or []
        # KEPADA format sesuai docx yang dipilih atasan langsung (varian), bukan nama bawahan
        # jadi kosongkan kepada_custom agar _build pakai var_info["kepada"] (YTH. DANSATSIBER/WADAN/ASINTELOPS/ASMINLOG)
        new_data['kepada_custom'] = ""
        # generate docx — template & KEPADA ikut varian pilihan atasan langsung
        doc = _build_disposisi_docx(new_data)
        fname_base = f"Lembar_Disposisi_{(new_data.get('no_agenda') or new_data.get('nomor_surat') or 'Satsiber').replace('/','-')}_{datetime.datetime.now().strftime('%Y%m%d%H%M%S')}_terusan"
        fname = fname_base + ".docx"
        fpath = os.path.join(DISPOSISI_UPLOAD_FOLDER, fname)
        doc.save(fpath)
        # — Lampiran PDF surat: file baru yang diupload ikut diteruskan,
        # kalau tidak ada file baru maka lampiran surat asal ikut terbawa —
        lampiran_simpan = orig.get('lampiran_nama_simpan') or ''
        lampiran_asli = orig.get('lampiran_nama_asli') or ''
        lampiran_file = request.files.get('lampiran_pdf') or request.files.get('lampiran') or request.files.get('file')
        if lampiran_file and lampiran_file.filename:
            nama_asli = lampiran_file.filename
            if '.' not in nama_asli or nama_asli.rsplit('.', 1)[1].lower() != 'pdf':
                return jsonify({"error": "Lampiran harus berformat PDF."}), 400
            lampiran_file.seek(0, os.SEEK_END)
            ukuran_mb = lampiran_file.tell() / (1024 * 1024)
            lampiran_file.seek(0)
            if ukuran_mb > 10:
                return jsonify({"error": "Ukuran PDF lampiran maksimal 10 MB."}), 400
            lampiran_simpan = f"lampiran_{datetime.datetime.now().strftime('%Y%m%d%H%M%S')}_{secure_filename(nama_asli)}"
            lampiran_file.save(os.path.join(DISPOSISI_UPLOAD_FOLDER, lampiran_simpan))
            lampiran_asli = nama_asli
        pengirim_nama = session.get('user_name') or (fdb.collection('users').document(saya).get().to_dict() or {}).get('fullname','')
        _public = get_cloudflared_url()
        _link = f'{_public}/USER/alat-bantuan.html' if _public else ''
        _lamp_path = os.path.join(DISPOSISI_UPLOAD_FOLDER, lampiran_simpan) if lampiran_simpan else ''
        ids_terusan, nama_terusan = [], []
        tg_terkirim = 0
        for tujuan_id, tujuan in daftar_tujuan:
            ref = fdb.collection('disposisi_surat').document()
            ref.set({
                'pengirim_id': saya,
                'pengirim_nama': pengirim_nama,
                'atasan_id': tujuan_id,
                'atasan_nama': tujuan.get('fullname') or '',
                'atasan_jabatan': tujuan.get('jabatan_atasan') or '',
                'varian': varian,
                'kepada_custom': "",  # sesuai format docx atasan langsung
                'perihal': new_data.get('perihal') or '',
                'nomor_surat': new_data.get('nomor_surat') or '',
                'no_agenda': new_data.get('no_agenda') or '',
                'surat_dari': new_data.get('surat_dari') or '',
                'jenis_surat': new_data.get('jenis_surat') or '',
                'tanggal_surat': new_data.get('tanggal_surat') or '',
                'tanggal_terima': new_data.get('tanggal_terima') or '',
                'diterima_dari': new_data.get('diterima_dari') or '',
                'diteruskan_kepada': new_data.get('diteruskan_kepada') or [],
                'isi_disposisi': new_data.get('isi_disposisi') or [],
                'catatan': new_data.get('catatan') or '',
                'paraf': new_data.get('paraf') or '',
                'paraf_image': new_data.get('paraf_image') or '',
                'asal_surat_id': surat_id,
                'file_nama_simpan': fname,
                'file_nama_asli': fname,
                'lampiran_nama_simpan': lampiran_simpan,
                'lampiran_nama_asli': lampiran_asli,
                'created_at': datetime.datetime.now().strftime('%d %b %Y %H:%M'),
                'created_ts': SERVER_TIMESTAMP,
                'status': 'baru',
                'dibaca_oleh': []
            })
            ids_terusan.append(ref.id)
            nama_terusan.append(tujuan.get('fullname') or '')
            try:
                if notifikasi_disposisi_telegram(tujuan, {
                        'perihal': new_data.get('perihal') or '',
                        'nomor_surat': new_data.get('nomor_surat') or '',
                        'no_agenda': new_data.get('no_agenda') or '',
                        'jenis_surat': new_data.get('jenis_surat') or '',
                        'pengirim_nama': pengirim_nama,
                        'file_docx': fpath,
                        'lampiran_pdf': _lamp_path,
                        'app_link': _link}):
                    tg_terkirim += 1
            except Exception as _e:
                print(f'[TELEGRAM] notifikasi terusan gagal: {_e}')
        if len(ids_terusan) == 1:
            _msg = f"Surat berhasil diteruskan ke {nama_terusan[0]} ({daftar_tujuan[0][1].get('jabatan_atasan') or daftar_tujuan[0][1].get('bagian') or ''})."
            if tg_terkirim:
                _msg += " Notifikasi Telegram terkirim."
            return jsonify({"message": _msg, "id": ids_terusan[0], "telegram_terkirim": tg_terkirim}), 201
        _msg = f"Surat berhasil diteruskan ke {len(ids_terusan)} tujuan: {', '.join(nama_terusan)}."
        if tg_terkirim:
            _msg += f" Notifikasi Telegram terkirim ke {tg_terkirim} penerima."
        return jsonify({"message": _msg, "ids": ids_terusan, "telegram_terkirim": tg_terkirim}), 201
    except Exception as e:
        import traceback; traceback.print_exc()
        return jsonify({"error": f"Gagal meneruskan: {str(e)}"}), 500

@app.route('/api/disposisi/surat-masuk/<string:surat_id>', methods=['DELETE'])
@login_required
@disposisi_required
def hapus_surat_masuk(surat_id):
    # Pantau-only (atasan 1-2 bawaan) tidak boleh menghapus
    if _hanya_pantau():
        return jsonify({"error": "Akun ini hanya dapat memantau surat."}), 403
    doc_ref = fdb.collection('disposisi_surat').document(surat_id)
    doc = doc_ref.get()
    if not doc.exists:
        return jsonify({"error": "Surat tidak ditemukan."}), 404
    x = doc.to_dict()
    saya = session.get('user_id')
    peran = session.get('user_role')
    # pemegang/pengirim/admin, atau staf persuratan (kelola arsip disposisi)
    boleh = (x.get('atasan_id') == saya or x.get('pengirim_id') == saya or peran == 'admin' or is_persuratan(saya))
    if not boleh:
        return jsonify({"error": "Anda tidak memiliki akses untuk menghapus."}), 403
    ok_pin, err_pin = _cek_pin_disposisi()
    if not ok_pin:
        return err_pin
    try:
        # Hapus satu rantai penuh (surat ini + seluruh terusannya di atasan)
        # agar riwayat atasan ikut menyesuaikan seperti riwayat persuratan.
        _target = _subtree_ids(surat_id) or [surat_id]
        _berkas = []
        for _sid in _target:
            try:
                _ddoc = fdb.collection('disposisi_surat').document(_sid).get()
                if _ddoc.exists:
                    _dx = _ddoc.to_dict() or {}
                    _berkas += [_dx.get('file_nama_simpan'), _dx.get('lampiran_nama_simpan')]
            except Exception:
                pass
        for _sid in _target:
            try:
                fdb.collection('disposisi_surat').document(_sid).delete()
            except Exception:
                pass
            try:
                for _lg in fdb.collection('disposisi_log').where('surat_id', '==', _sid).stream():
                    try:
                        _lg.reference.delete()
                    except Exception:
                        pass
            except Exception:
                pass
        # Hapus file fisik hanya bila tak ada surat lain yang masih memakainya
        # (satu berkas bisa dipakai bersama beberapa salinan terusan).
        for _fn in {f for f in _berkas if f}:
            try:
                _masih = False
                for _kol in ('file_nama_simpan', 'lampiran_nama_simpan'):
                    try:
                        if next(fdb.collection('disposisi_surat').where(_kol, '==', _fn).limit(1).stream(), None):
                            _masih = True
                            break
                    except Exception:
                        pass
                if not _masih:
                    _fp = os.path.join(DISPOSISI_UPLOAD_FOLDER, _fn)
                    if os.path.exists(_fp):
                        os.remove(_fp)
            except OSError:
                pass
        _n_ikut = len(_target) - 1
        _pesan = "Surat berhasil dihapus."
        if _n_ikut > 0:
            _pesan += f" {_n_ikut} surat terusan di atasan ikut terhapus."
        return jsonify({"message": _pesan, "terusan_terhapus": _n_ikut}), 200
    except Exception as e:
        return jsonify({"error": f"Gagal menghapus: {str(e)}"}), 500

@app.route('/api/disposisi/surat-masuk/<string:surat_id>/edit', methods=['POST'])
@login_required
@disposisi_required
def edit_surat_masuk(surat_id):
    """Ubah isi surat disposisi: perihal, nomor, catatan, paraf, varian,
    dan centangan (diteruskan_kepada / isi_disposisi). Tujuan/penerima TIDAK
    diubah di sini (pakai Teruskan). Boleh: admin, staf persuratan, atau
    pemegang/pengirim surat. Pantau-only (atasan 1-2): tidak. Perubahan
    dicatat di riwayat_ubah dan tampil di Jejak."""
    try:
        if _hanya_pantau():
            return jsonify({"error": "Akun ini hanya dapat memantau surat."}), 403
        doc_ref = fdb.collection('disposisi_surat').document(surat_id)
        doc = doc_ref.get()
        if not doc.exists:
            return jsonify({"error": "Surat tidak ditemukan."}), 404
        x = doc.to_dict()
        saya = session.get('user_id')
        peran = session.get('user_role')
        if not (x.get('atasan_id') == saya or x.get('pengirim_id') == saya or peran == 'admin' or is_persuratan(saya)):
            return jsonify({"error": "Anda tidak memiliki akses untuk mengubah surat ini."}), 403
        ok_pin, err_pin = _cek_pin_disposisi()
        if not ok_pin:
            return err_pin
        data = request.get_json(force=True, silent=True) or {}
        patch = {}
        for _f in ('perihal', 'nomor_surat', 'catatan', 'paraf', 'varian'):
            if isinstance(data.get(_f), str):
                patch[_f] = data[_f].strip()[:500]
        for _f in ('diteruskan_kepada', 'isi_disposisi'):
            if isinstance(data.get(_f), list):
                patch[_f] = [str(_v)[:80] for _v in data[_f][:20]]
        if not patch:
            return jsonify({"error": "Tidak ada perubahan yang dikirim."}), 400
        nama = session.get('user_name') or _nama_pengguna(saya)
        waktu = datetime.datetime.now().strftime('%d %b %Y %H:%M')
        _log = list(x.get('riwayat_ubah') or [])
        _log.append({'oleh': nama, 'pada': waktu,
                     'diubah': sorted([_k for _k in patch if _k not in ('diubah_pada', 'diubah_oleh')])})
        patch.update({'diubah_pada': waktu, 'diubah_oleh': saya, 'diubah_oleh_nama': nama,
                      'riwayat_ubah': _log[-20:]})
        doc_ref.update(patch)
        # — Rambatkan ubahan ke surat terusan (salinan di atasan) agar
        # riwayat atasan sama dengan persuratan. Hanya kolom yang nilainya
        # masih sama dengan sebelum diubah yang diselaraskan (kolom yang
        # sudah dikustom di hilir tidak ditimpa).
        _isi = {k: v for k, v in patch.items()
                if k in ('perihal', 'nomor_surat', 'catatan', 'paraf', 'varian',
                         'diteruskan_kepada', 'isi_disposisi')}
        _selaras = 0
        if _isi:
            try:
                for _sid in _subtree_ids(surat_id):
                    if _sid == surat_id:
                        continue
                    try:
                        _cref = fdb.collection('disposisi_surat').document(_sid)
                        _cdoc = _cref.get()
                        if not _cdoc.exists:
                            continue
                        _cx = _cdoc.to_dict() or {}
                        _pp = {k: v for k, v in _isi.items() if _cx.get(k) == x.get(k)}
                        if not _pp:
                            continue
                        _clog = list(_cx.get('riwayat_ubah') or [])
                        _clog.append({'oleh': f'{nama} (selaras hulu)', 'pada': waktu,
                                      'diubah': sorted(_pp.keys())})
                        _pp.update({'diubah_pada': waktu, 'diubah_oleh': saya,
                                    'diubah_oleh_nama': nama, 'riwayat_ubah': _clog[-20:]})
                        _cref.update(_pp)
                        _selaras += 1
                    except Exception:
                        continue
            except Exception:
                pass
        _pesan = "Surat berhasil diubah."
        if _selaras:
            _pesan += f" Diselaraskan ke {_selaras} surat terusan di atasan."
        return jsonify({"message": _pesan, "diubah_pada": waktu, "diselaraskan": _selaras}), 200
    except Exception as e:
        return jsonify({"error": str(e)}), 500


# ===================== RIWAYAT SURAT DISPOSISI =====================
# Akses ketat: hanya yang terlibat (pengirim/penerima) atau admin.
# Atasan TIDAK dapat mengintip rantai timnya.

def _catat_log_disposisi(surat_id, aksi):
    """Catat aktivitas (unduh/lampiran) ke disposisi_log. Gagal catat tak menggagalkan aksi."""
    try:
        saya = session.get('user_id')
        nama = session.get('user_name') or ''
        if not nama and saya:
            _u = fdb.collection('users').document(saya).get()
            if _u.exists:
                nama = (_u.to_dict() or {}).get('fullname') or ''
        fdb.collection('disposisi_log').document().set({
            'surat_id': surat_id,
            'aktor_id': saya,
            'aktor_nama': nama,
            'aksi': aksi,
            'waktu': datetime.datetime.now().strftime('%d %b %Y %H:%M'),
            'waktu_ts': SERVER_TIMESTAMP
        })
    except Exception as _e:
        print(f'[DISPOSISI-LOG] gagal catat: {_e}')


def _nama_pengguna(uid):
    try:
        _d = fdb.collection('users').document(uid).get()
        if _d.exists:
            return (_d.to_dict() or {}).get('fullname') or uid
    except Exception:
        pass
    return uid or '-'


def _rantai_node(surat_id, batas=50):
    """Ambil rantai penuh: telusur ke akar via asal_surat_id + semua turunan (BFS)."""
    semua = {}
    # ke hulu
    cur = surat_id
    while cur and len(semua) < batas:
        _d = fdb.collection('disposisi_surat').document(cur).get()
        if not _d.exists:
            break
        semua[cur] = _d.to_dict()
        cur = (_d.to_dict() or {}).get('asal_surat_id')
    # ke hilir (BFS dari akar)
    akar = cur or surat_id
    if akar not in semua:
        _d = fdb.collection('disposisi_surat').document(akar).get()
        if _d.exists:
            semua[akar] = _d.to_dict()
    antre = [akar]
    while antre and len(semua) < batas:
        _pid = antre.pop(0)
        for _c in fdb.collection('disposisi_surat').where('asal_surat_id', '==', _pid).stream():
            if _c.id not in semua:
                semua[_c.id] = _c.to_dict()
                antre.append(_c.id)
    return semua


def _terlibat(saya, peran, node):
    if peran == 'admin':
        return True
    return node.get('pengirim_id') == saya or node.get('atasan_id') == saya


def _subtree_ids(surat_id, batas=100):
    """Kumpulkan ID surat ini + seluruh terusannya (hilir via asal_surat_id, BFS).

    Dipakai agar Ubah/Hapus oleh persuratan merambat ke salinan terusan di
    atasan — riwayat atasan & persuratan tetap sama."""
    kumpul = []
    antre = [surat_id]
    lihat = set()
    while antre and len(kumpul) < batas:
        _pid = antre.pop(0)
        if _pid in lihat:
            continue
        lihat.add(_pid)
        try:
            _d = fdb.collection('disposisi_surat').document(_pid).get()
            if not _d.exists:
                continue
        except Exception:
            continue
        kumpul.append(_pid)
        try:
            for _c in fdb.collection('disposisi_surat').where('asal_surat_id', '==', _pid).stream():
                if _c.id not in lihat:
                    antre.append(_c.id)
        except Exception:
            pass
    return kumpul


@app.route('/api/disposisi/riwayat-keluar', methods=['GET'])
@login_required
@disposisi_required
def riwayat_keluar_disposisi():
    """Surat yang pernah SAYA kirim/teruskan. Admin boleh filter ?pengirim=<id>."""
    try:
        saya = session.get('user_id')
        peran = session.get('user_role')
        target = request.args.get('pengirim') or ''
        if peran == 'admin' and target:
            docs = list(fdb.collection('disposisi_surat').where('pengirim_id', '==', target).stream())
        elif peran == 'admin' and not target:
            docs = list(fdb.collection('disposisi_surat').stream())
        else:
            docs = list(fdb.collection('disposisi_surat').where('pengirim_id', '==', saya).stream())
        hasil = _serialize_surat_masuk(docs)
        # tambah info keterbacaan lawan + jumlah terusan
        for h in hasil:
            try:
                _d = fdb.collection('disposisi_surat').document(h['id']).get().to_dict() or {}
                h['dibaca_penerima'] = bool((_d.get('dibaca_oleh') or []) and
                                            any(p != saya for p in (_d.get('dibaca_oleh') or [])))
                h['jumlah_terusan'] = sum(1 for _c in fdb.collection('disposisi_surat').where('asal_surat_id', '==', h['id']).stream())
            except Exception:
                h['dibaca_penerima'] = False
                h['jumlah_terusan'] = 0
        return jsonify(hasil)
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route('/api/disposisi/rantai/<string:surat_id>', methods=['GET'])
@login_required
@disposisi_required
def rantai_disposisi(surat_id):
    """Rantai terusan penuh sebuah surat. Hanya yang terlibat atau admin."""
    try:
        saya = session.get('user_id')
        peran = session.get('user_role')
        semua = _rantai_node(surat_id)
        if not semua:
            return jsonify({"error": "Surat tidak ditemukan."}), 404
        if not any(_terlibat(saya, peran, n) for n in semua.values()):
            return jsonify({"error": "Anda tidak terlibat dalam rantai surat ini."}), 403
        # buat objek pseudo-doc agar bisa pakai serializer (dengan sudah_dibaca)
        class _D:
            def __init__(self, i, d):
                self.id = i
                self._d = d
            def to_dict(self):
                return self._d
        urut = sorted(semua.items(), key=lambda kv: kv[1].get('created_at') or '')
        hasil = _serialize_surat_masuk([_D(i, d) for i, d in urut])
        for h in hasil:
            h['asal_surat_id'] = (semua.get(h['id']) or {}).get('asal_surat_id')
        akar = next((i for i, d in semua.items() if not d.get('asal_surat_id')), surat_id)
        return jsonify({"akar_id": akar, "rantai": hasil})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route('/api/disposisi/surat-masuk/<string:surat_id>/jejak', methods=['GET'])
@login_required
@disposisi_required
def jejak_disposisi(surat_id):
    """Timeline satu surat: dibuat, dibaca, diunduh, diteruskan ke mana."""
    try:
        saya = session.get('user_id')
        peran = session.get('user_role')
        doc = fdb.collection('disposisi_surat').document(surat_id).get()
        if not doc.exists:
            return jsonify({"error": "Surat tidak ditemukan."}), 404
        x = doc.to_dict()
        if not _terlibat(saya, peran, x):
            return jsonify({"error": "Anda tidak terlibat dalam surat ini."}), 403
        peristiwa = [{
            'waktu': x.get('created_at') or '',
            'teks': f"Dibuat & dikirim oleh {x.get('pengirim_nama') or '-'} ke {x.get('atasan_nama') or '-'}",
            'jenis': 'kirim'
        }]
        for _pid in (x.get('dibaca_oleh') or []):
            peristiwa.append({'waktu': '', 'teks': f"Dibaca oleh {_nama_pengguna(_pid)}", 'jenis': 'baca'})
        if x.get('konfirmasi_oleh'):
            peristiwa.append({'waktu': x.get('konfirmasi_pada') or '',
                              'teks': f"Dikonfirmasi diterima oleh {x.get('konfirmasi_oleh_nama') or '-'}",
                              'jenis': 'konfirmasi'})
        for _ub in (x.get('riwayat_ubah') or []):
            peristiwa.append({'waktu': _ub.get('pada') or '',
                              'teks': f"Diubah ({', '.join(_ub.get('diubah') or []) or 'isi'}) oleh {_ub.get('oleh') or '-'}",
                              'jenis': 'ubah'})
        for _lg in fdb.collection('disposisi_log').where('surat_id', '==', surat_id).stream():
            _l = _lg.to_dict() or {}
            _aksi = {'unduh': 'Mengunduh berkas', 'lampiran': 'Membuka lampiran PDF', 'konfirmasi': 'Konfirmasi diterima'}.get(_l.get('aksi'), _l.get('aksi'))
            _jenis = 'konfirmasi' if _l.get('aksi') == 'konfirmasi' else 'unduh'
            peristiwa.append({'waktu': _l.get('waktu') or '',
                              'teks': f"{_aksi} — {_l.get('aktor_nama') or '-'}", 'jenis': _jenis})
        anak = []
        for _c in fdb.collection('disposisi_surat').where('asal_surat_id', '==', surat_id).stream():
            _cd = _c.to_dict() or {}
            anak.append({'id': _c.id, 'ke': _cd.get('atasan_nama') or '-',
                         'oleh': _cd.get('pengirim_nama') or '-',
                         'waktu': _cd.get('created_at') or ''})
            peristiwa.append({'waktu': _cd.get('created_at') or '',
                              'teks': f"Diteruskan oleh {_cd.get('pengirim_nama') or '-'} ke {_cd.get('atasan_nama') or '-'}",
                              'jenis': 'teruskan'})
        return jsonify({"perihal": x.get('perihal'), "peristiwa": peristiwa, "terusan": anak})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route('/api/disposisi/surat-masuk/<string:surat_id>/konfirmasi', methods=['POST'])
@login_required
@disposisi_required
def konfirmasi_disposisi(surat_id):
    """Konfirmasi eksplisit 'sudah diterima' oleh PEMEGANG surat.
    Idempoten: konfirmasi ulang tetap sukses. Tercatat di jejak."""
    try:
        saya = session.get('user_id')
        peran = session.get('user_role')
        doc_ref = fdb.collection('disposisi_surat').document(surat_id)
        doc = doc_ref.get()
        if not doc.exists:
            return jsonify({"error": "Surat tidak ditemukan."}), 404
        x = doc.to_dict()
        if x.get('atasan_id') != saya and peran != 'admin':
            return jsonify({"error": "Hanya pemegang surat yang dapat mengonfirmasi."}), 403
        if x.get('konfirmasi_oleh'):
            return jsonify({"message": "Sudah dikonfirmasi sebelumnya.",
                            "konfirmasi_oleh": x.get('konfirmasi_oleh_nama'),
                            "konfirmasi_pada": x.get('konfirmasi_pada')}), 200
        nama = session.get('user_name') or _nama_pengguna(saya)
        waktu = datetime.datetime.now().strftime('%d %b %Y %H:%M')
        doc_ref.update({'konfirmasi_oleh': saya, 'konfirmasi_oleh_nama': nama,
                        'konfirmasi_pada': waktu, 'status': 'dikonfirmasi'})
        return jsonify({"message": f"Surat dikonfirmasi diterima oleh {nama}.",
                        "konfirmasi_oleh": nama, "konfirmasi_pada": waktu}), 200
    except Exception as e:
        return jsonify({"error": str(e)}), 500
# ===================================================================================

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True, threaded=True)
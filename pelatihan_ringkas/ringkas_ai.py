"""AI Ringkas Dokumen Bahasa Indonesia (extractive, khusus dokumen dinas).

Arsitektur:
  teks -> pra-pembersihan -> splitter kalimat ID -> TF-IDF (Sastrawi)
  -> PageRank kemiripan + fitur posisi/panjang/kata-kunci -> pilih kalimat terbaik.

Bobot disimpan di ringkas_config_v1.json (hasil tuning lawan korpus).
"""
import json
import math
import os
import re

import numpy as np

_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(_DIR, 'ringkas_config_v1.json')

# ---------------- Konfigurasi default (ditimpa file config bila ada) ----------------
DEFAULT_CONFIG = {
    "bobot_pagerank": 0.55,
    "bobot_posisi": 0.15,
    "bobot_panjang": 0.10,
    "bobot_kata_kunci": 0.12,
    "bobot_angka": 0.08,
    "ambang_kemiripan": 0.10,
    "damping": 0.85,
    "kata_kunci": ["menimbang", "mengingat", "memperhatikan", "memutuskan", "menetapkan",
                   "disposisi", "instruksi", "evaluasi", "kesimpulan", "hasil", "saran",
                   "tindak lanjut", "lapor", "penting", "segera", "wajib", "sasaran"],
}


def muat_config():
    cfg = dict(DEFAULT_CONFIG)
    try:
        with open(CONFIG_PATH, encoding='utf-8') as fh:
            cfg.update(json.load(fh))
    except (OSError, ValueError):
        pass
    return cfg


# ---------------- Stemmer/stopword Sastrawi (dibuat sekali) ----------------
_STEMMER = None
_STOPWORDS = None


def _sastrawi():
    global _STEMMER, _STOPWORDS
    if _STEMMER is None:
        from Sastrawi.Stemmer.StemmerFactory import StemmerFactory
        from Sastrawi.StopWordRemover.StopWordRemoverFactory import StopWordRemoverFactory
        _STEMMER = StemmerFactory().create_stemmer()
        _STOPWORDS = set(StopWordRemoverFactory().get_stop_words())
    return _STEMMER, _STOPWORDS


# ---------------- 1. Pra-pembersihan ----------------
_KOP_BERULANG_MIN = 2


def bersihkan_teks(text):
    text = text.replace('\r\n', '\n').replace('\r', '\n')
    # sambung kata terpotong hyphen di akhir baris: "pemantau-\nan" -> "pemantauan"
    text = re.sub(r'(\w)-\n(\w)', r'\1\2', text)
    baris = [b.strip() for b in text.split('\n')]
    # buang baris nomor halaman murni & baris kosong
    baris = [b for b in baris if b and not re.fullmatch(r'\d{1,4}', b)]
    # buang kop/slogan yang berulang (>=2x kemunculan persis)
    hitung = {}
    for b in baris:
        hitung[b] = hitung.get(b, 0) + 1
    baris = [b for b in baris if hitung[b] < _KOP_BERULANG_MIN or len(b.split()) > 6]
    text = '\n'.join(baris)
    text = re.sub(r'[ \t]+', ' ', text)
    return text.strip()


# ---------------- 2. Splitter kalimat Bahasa Indonesia ----------------
# Singkatan dinas/gelar yang titiknya BUKAN akhir kalimat.
SINGKATAN = {
    'no', 'bpk', 'ibu', 'sdr', 'sdri', 'dr', 'drs', 'ir', 'tgl', 's.d', 'd.a',
    'dsb', 'dll', 'dst', 'cth', 'mis', 'sbb', 'tsb', 'ybs', 'yth',
    'a.n', 'u.b', 'u.p', 'dpp', 'st', 's.t', 's.e', 's.h', 's.ip', 's.si',
    'm.m', 'm.si', 's.pd', 's.sos', 's.psi', 'a.md', 's.st', 's.tr', 'brigadir',
    'jan', 'feb', 'mar', 'apr', 'mei', 'jun', 'jul', 'agu', 'sep', 'okt',
    'nov', 'des', 'pkl', 'wib', 'wita', 'wit', 'rp', 'tlp',
    'telp', 'hp', 'jl', 'jln', 'rt', 'rw', 'kel', 'kec', 'kab', 'prov',
    'kol', 'letkol', 'mayor', 'kapten', 'letnan', 'serka', 'serda', 'kopral',
    'jenderal', 'brigjen', 'irjen', 'komjen',
}
_PLACEHOLDER = '\ue000'  # char privat, tidak muncul di teks normal


def _lindungi_singkatan(text):
    def ganti(m):
        kata = m.group(1).lower().strip('.').lower()
        frag = m.group(0)
        # cek token terakhir sebelum titik termasuk singkatan?
        token = re.split(r'\s+', frag.strip())[-1].rstrip('.').lower()
        if token in SINGKATAN or re.fullmatch(r'[a-z]\.[a-z]\.?', token):
            return frag.replace('.', _PLACEHOLDER)
        return frag
    # lindungi pola "X." di mana X singkatan ATAU inisial "A." / "A.B."
    text = re.sub(r'(\b[A-Za-z](?:\.[A-Za-z])?\.)', ganti, text)
    # lindungi angka desimal & penomoran jam "10.00", tanggal "25.09"
    text = re.sub(r'(\d)\.(\d)', rf'\1{_PLACEHOLDER}\2', text)
    return text


def split_kalimat_id(text):
    """Pecah teks menjadi daftar kalimat utuh (tidak sepotong kata)."""
    if not text or not text.strip():
        return []
    t = _lindungi_singkatan(text)
    # Pecah pada baris bernomor "1. ...", "2. ..." sebagai kalimat sendiri
    t = re.sub(r'\n\s*(\d{1,3}[.)])\s+', r'\n\1 ', t)
    kalimat = []
    for blok in t.split('\n'):
        blok = blok.strip()
        if not blok:
            continue
        # belah pada . ! ? yang diikuti spasi + huruf kapital/angka
        bagian = re.split(r'(?<=[.!?])\s+(?=[A-Z0-9"\u201c])', blok)
        for b in bagian:
            b = b.strip().rstrip(';').strip()
            if not b:
                continue
            b = b.replace(_PLACEHOLDER, '.')
            # buang fragmen terlalu pendek yang BUKAN penomoran (penyebab "sepotong kata")
            n_kata = len(b.split())
            if n_kata < 3 and not re.match(r'^\d{1,3}[.)]', b):
                if kalimat:
                    kalimat[-1] = kalimat[-1] + ' ' + b
                    continue
                # fragmen pembuka sendirian: gabung ke kalimat berikut bila ada
            kalimat.append(b)
    # gabung fragmen pembuka pendek ke kalimat berikutnya
    digabung = []
    i = 0
    while i < len(kalimat):
        k = kalimat[i]
        if len(k.split()) < 3 and i + 1 < len(kalimat) and not re.match(r'^\d{1,3}[.)]', k):
            digabung.append(k + ' ' + kalimat[i + 1])
            i += 2
        else:
            digabung.append(k)
            i += 1
    return [k for k in digabung if len(k.split()) >= 2]


# ---------------- 3. Ekstraksi docx lengkap ----------------
def ekstrak_docx(path):
    from docx import Document
    doc = Document(path)
    teks = []
    for p in doc.paragraphs:
        if p.text.strip():
            teks.append(p.text.strip())
    for tbl in doc.tables:
        for row in tbl.rows:
            sel = [c.text.strip() for c in row.cells if c.text.strip()]
            if sel:
                teks.append(' | '.join(sel))
    # header & footer
    for sec in doc.sections:
        for part in (sec.header, sec.footer):
            for p in part.paragraphs:
                if p.text.strip():
                    teks.append(p.text.strip())
    return '\n'.join(teks)


# ---------------- 4. Skor hibrida ----------------
def _tokenize(kalimat, stemmer, stopwords):
    kata = re.findall(r'[a-zA-Z]+', kalimat.lower())
    return [stemmer.stem(w) for w in kata if w not in stopwords and len(w) > 2]


def _tfidf_matrix(token_docs):
    # token_docs: list[list[str]]
    df = {}
    for toks in token_docs:
        for w in set(toks):
            df[w] = df.get(w, 0) + 1
    n = max(len(token_docs), 1)
    kosakata = {w: i for i, w in enumerate(df)}
    mat = np.zeros((len(token_docs), len(kosakata)))
    for i, toks in enumerate(token_docs):
        total = len(toks) or 1
        tf = {}
        for w in toks:
            tf[w] = tf.get(w, 0) + 1
        for w, c in tf.items():
            mat[i, kosakata[w]] = (c / total) * math.log(n / df[w])
    return mat


def _pagerank(sim, damping=0.85, iters=60):
    n = sim.shape[0]
    if n == 1:
        return np.ones(1)
    jumlah = sim.sum(axis=1, keepdims=True)
    jumlah[jumlah == 0] = 1.0
    trans = sim / jumlah
    skor = np.full(n, 1.0 / n)
    for _ in range(iters):
        baru = (1 - damping) / n + damping * trans.T.dot(skor)
        if np.abs(baru - skor).sum() < 1e-6:
            break
        skor = baru
    return skor


def skor_kalimat(kalimat_list, cfg=None, token_docs=None):
    cfg = cfg or muat_config()
    if token_docs is None:
        stemmer, stopwords = _sastrawi()
        token_docs = [_tokenize(k, stemmer, stopwords) for k in kalimat_list]
    n = len(kalimat_list)
    mat = _tfidf_matrix(token_docs)
    # matriks kemiripan kosinus
    norm = np.linalg.norm(mat, axis=1, keepdims=True)
    norm[norm == 0] = 1.0
    sim = (mat / norm) @ (mat / norm).T
    np.fill_diagonal(sim, 0.0)
    sim[sim < cfg['ambang_kemiripan']] = 0.0
    pr = _pagerank(sim, cfg['damping'])
    if pr.max() > 0:
        pr = pr / pr.max()

    kunci = cfg['kata_kunci']
    hasil = []
    for i, k in enumerate(kalimat_list):
        nk = len(k.split())
        # posisi: kalimat awal dokumen dinas memuat inti
        posisi = 1.0 - (i / max(n, 1)) * 0.7
        # panjang ideal 8-35 kata
        if nk < 4:
            panjang = nk / 4.0 * 0.4
        elif nk <= 35:
            panjang = 1.0
        else:
            panjang = max(0.3, 35.0 / nk)
        kl = k.lower()
        kk = sum(1 for w in kunci if w in kl) / 3.0
        kk = min(kk, 1.0)
        angka = 1.0 if re.search(r'\d', k) else 0.0
        skor = (cfg['bobot_pagerank'] * pr[i] + cfg['bobot_posisi'] * posisi
                + cfg['bobot_panjang'] * panjang + cfg['bobot_kata_kunci'] * kk
                + cfg['bobot_angka'] * angka)
        hasil.append(skor)
    return hasil


MODES = {
    'singkat': {'rasio': 0.15, 'maks': 5},
    'sedang': {'rasio': 0.30, 'maks': 10},
    'lengkap': {'rasio': 0.50, 'maks': 20},
}


def ringkas(teks, mode='sedang', cfg=None, min_kalimat=2, jumlah=None):
    """Kembalikan (ringkasan, statistik). `jumlah` menimpa aturan mode bila diisi."""
    cfg = cfg or muat_config()
    bersih = bersihkan_teks(teks)
    kalimat = split_kalimat_id(bersih)
    total = len(kalimat)
    if total == 0:
        return "Dokumen tidak berisi teks untuk diringkas.", {'total': 0, 'diambil': 0, 'kompresi': 0}
    if total <= min_kalimat:
        return bersih.strip(), {'total': total, 'diambil': total, 'kompresi': 0}
    if jumlah is None:
        m = MODES.get(mode, MODES['sedang'])
        jumlah = max(min_kalimat, min(m['maks'], round(total * m['rasio'])))
    else:
        jumlah = max(min_kalimat, min(total, int(jumlah)))
    skor = skor_kalimat(kalimat, cfg)
    peringkat = sorted(range(total), key=lambda i: skor[i], reverse=True)[:jumlah]
    terpilih = sorted(peringkat)  # urutan baca asli
    hasil = '\n\n'.join(kalimat[i] for i in terpilih)
    return hasil, {'total': total, 'diambil': len(terpilih),
                   'kompresi': round(100 * (1 - len(terpilih) / total))}


# ---------------- 5. Penjelasan terstruktur ----------------
# Mengelompokkan kalimat terpilih ke tiga peran (tanpa mengarang):
#   deskripsi  : gambaran umum isi dokumen (kalimat sentral/generik).
#   contoh     : hal konkret — angka/tanggal/nama atau penanda contoh.
#   penjelasan : kalimat sebab-akibat/tujuan + keterkaitan topik.
CONTOH_POLAS = [
    r'\bcontoh\b', r'\bmisal\b', r'misalnya', r'\bseperti\b', r'umpama',
    r'ibarat', r'permisalan', r'\byaitu\b', r'\byakni\b', r'antara lain',
    r'sebagai berikut', r'\d',
]
PENJELAS_POLAS = [
    r'\bkarena\b', r'\bsebab\b', r'sehingga', r'\bagar\b', r'\buntuk\b',
    r'\bmaka\b', r'artinya', r'dengan demikian', r'oleh karena itu',
    r'diharapkan', r'bertujuan', r'disebabkan', r'menyebabkan',
    r'dengan kata lain', r'hal ini menunjukkan', r'hal tersebut',
    r'dipandang perlu', r'dengan ini',
]
DESKRIPSI_POLAS = [
    r'\badalah\b', r'merupakan', r'tentang', r'membahas', r'berisi\b',
    r'dokumen ini', r'surat ini', r'memo ini', r'laporan ini', r'disposisi ini',
]


def _cocok(kalimat, polas):
    kl = kalimat.lower()
    return [p.strip('\\b') for p in polas if re.search(p, kl)]


def struktur_penjelasan(kalimat_terpilih, skor_terpilih=None, cfg=None, maks_per_bagian=4):
    """Kelompokkan kalimat terpilih menjadi deskripsi/contoh/penjelasan.

    Kembalikan dict {deskripsi:[{kalimat,mengapa}], contoh:[...],
    penjelasan_relevan:[...]}. Setiap ember diisi minimal 1 bila ada
    kalimat (diambil dari sisa berperingkat bila penanda tak cukup).
    """
    cfg = cfg or muat_config()
    kunci = cfg.get('kata_kunci', [])
    n = len(kalimat_terpilih)
    urut = sorted(range(n),
                  key=lambda i: (skor_terpilih[i] if skor_terpilih else 0),
                  reverse=True)

    def alasan(i, peran):
        k = kalimat_terpilih[i]
        als = []
        if re.search(r'\d', k):
            als.append('data konkret')
        kk = [w for w in kunci if w in k.lower()][:2]
        if kk:
            als.append('terkait ' + ', '.join(kk))
        als.append(peran)
        return '; '.join(als)

    ember = {'deskripsi': [], 'contoh': [], 'penjelasan': []}
    dipakai = set()
    # putaran 1: berdasar penanda (contoh dulu, lalu penjelasan, lalu deskripsi)
    for i in urut:
        if len(ember['contoh']) >= maks_per_bagian:
            break
        k = kalimat_terpilih[i]
        if _cocok(k, CONTOH_POLAS):
            ember['contoh'].append({'kalimat': k, 'mengapa': alasan(i, 'contoh konkret')})
            dipakai.add(i)
    for i in urut:
        if i in dipakai:
            continue
        if len(ember['penjelasan']) >= maks_per_bagian:
            break
        k = kalimat_terpilih[i]
        if _cocok(k, PENJELAS_POLAS):
            ember['penjelasan'].append({'kalimat': k, 'mengapa': alasan(i, 'penjelasan sebab/tujuan')})
            dipakai.add(i)
    for i in urut:
        if i in dipakai:
            continue
        if len(ember['deskripsi']) >= maks_per_bagian:
            break
        k = kalimat_terpilih[i]
        if _cocok(k, DESKRIPSI_POLAS):
            ember['deskripsi'].append({'kalimat': k, 'mengapa': alasan(i, 'gambaran umum')})
            dipakai.add(i)
    # putaran 2: sisa berperingkat mengisi ember yang masih kosong
    # (prioritas deskripsi <- kalimat paling sentral lebih dulu)
    for i in urut:
        if i in dipakai:
            continue
        k = kalimat_terpilih[i]
        if len(ember['deskripsi']) < 1:
            ember['deskripsi'].append({'kalimat': k, 'mengapa': alasan(i, 'kalimat paling sentral')})
        elif len(ember['penjelasan']) < 1:
            ember['penjelasan'].append({'kalimat': k, 'mengapa': alasan(i, 'keterkaitan topik')})
        elif len(ember['contoh']) < 1:
            ember['contoh'].append({'kalimat': k, 'mengapa': alasan(i, 'pendukung konkret')})
        else:
            break
        dipakai.add(i)
    # putaran 3: lengkapi tiap ember sampai batas (sisa terbaik)
    for i in urut:
        if len(ember['deskripsi']) >= maks_per_bagian \
                and len(ember['contoh']) >= maks_per_bagian \
                and len(ember['penjelasan']) >= maks_per_bagian:
            break
        if i in dipakai:
            continue
        k = kalimat_terpilih[i]
        if _cocok(k, CONTOH_POLAS) and len(ember['contoh']) < maks_per_bagian:
            ember['contoh'].append({'kalimat': k, 'mengapa': alasan(i, 'contoh konkret')})
        elif _cocok(k, PENJELAS_POLAS) and len(ember['penjelasan']) < maks_per_bagian:
            ember['penjelasan'].append({'kalimat': k, 'mengapa': alasan(i, 'penjelasan sebab/tujuan')})
        elif len(ember['deskripsi']) < maks_per_bagian:
            ember['deskripsi'].append({'kalimat': k, 'mengapa': alasan(i, 'gambaran umum')})
        else:
            continue
        dipakai.add(i)
    return ember

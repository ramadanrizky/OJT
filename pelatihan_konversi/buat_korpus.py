"""Bangun korpus latih tata letak: 11 PDF berlabel + labels.json.
Jalankan:  python pelatihan_konversi\\buat_korpus.py
Hasil:     pelatihan_konversi\\korpus\\*.pdf + labels.json
"""
import json
import os

import fitz

DASAR = os.path.dirname(os.path.abspath(__file__))
KORPUS = os.path.join(DASAR, 'korpus')
os.makedirs(KORPUS, exist_ok=True)

LABELS = {"kasus": []}


def tulis_terbungkus(page, teks, x, y, batas, size=11, **kw):
    yy = y
    baris = ""
    for kata in teks.split():
        t = (baris + " " + kata).strip()
        if fitz.get_text_length(t, fontsize=size) > batas:
            page.insert_text((x, yy), baris, fontsize=size, **kw)
            yy += size + 3
            baris = kata
        else:
            baris = t
    page.insert_text((x, yy), baris, fontsize=size, **kw)
    return yy + size + 3


def tulis_tengah(page, teks, y, size=12, **kw):
    """Tulis teks benar-benar di tengah halaman (simetris presisi)."""
    W = page.rect.width
    x = (W - fitz.get_text_length(teks, fontsize=size)) / 2
    page.insert_text((x, y), teks, fontsize=size, **kw)
    return y + size + 4


def simpan(doc, nama, ekspektasi):
    path = os.path.join(KORPUS, nama)
    doc.save(path)
    doc.close()
    ekspektasi["pdf"] = nama
    LABELS["kasus"].append(ekspektasi)
    print("korpus:", nama)


# ---- 1. dasar: judul + justify + indent + warna + kanan + tabel + gambar ----
d = fitz.open()
p = d.new_page()
p.insert_text((200, 70), "SURAT KEPUTUSAN", fontsize=16)
p.insert_text((90, 110), "Menimbang :", fontsize=12)
t1 = ("bahwa untuk meningkatkan kesiapan operasional satuan dipandang perlu menetapkan "
      "kebijakan baru tentang pembinaan personel secara menyeluruh dan berkelanjutan di "
      "lingkungan satuan kerja masing-masing tanpa terkecuali bagi seluruh anggota.")
y = tulis_terbungkus(p, t1, 90, 130, 400)
y = tulis_terbungkus(p, "Ditetapkan di Jakarta pada tanggal 5 Oktober 2026 oleh pejabat berwenang.", 120, y + 8, 370)
y = tulis_terbungkus(p, "pada tanggal tersebut disaksikan para perwira dan seluruh anggota hadir.", 90, y + 2, 400)
p.insert_text((90, y + 6), "Sangat Segera", fontsize=12, color=(1, 0, 0))
p.insert_text((400, y + 30), "Ttd. Komandan", fontsize=11)
for r in range(3):
    for c in range(3):
        rc = fitz.Rect(90 + c * 130, y + 48 + r * 20, 90 + (c + 1) * 130, y + 48 + (r + 1) * 20)
        p.draw_rect(rc)
        p.insert_text((rc.x0 + 4, rc.y0 + 14),
                      ["Pos", "Status", "Ket"][c] if r == 0 else [f"Pos {r}", "Aman", "OK"][c], fontsize=11)
pix = fitz.Pixmap(fitz.csRGB, fitz.Rect(0, 0, 200, 100), 0)
pix.set_rect(pix.irect, (30, 144, 200))
p.insert_image(fitz.Rect(90, y + 115, 290, y + 165), pixmap=pix)
p.insert_text((270, 800), "3", fontsize=10)
simpan(d, "01-dasar.pdf", {
    "margin_L_in": 1.25,
    "paragraf": [
        {"kunci": "SURAT KEPUTUSAN", "align": "HEADING"},
        {"kunci": "kesiapan operasional", "align": "LEFT", "indent": False},
        {"kunci": "Ditetapkan di Jakarta", "align": "LEFT", "indent": True},
        {"kunci": "Sangat Segera", "align": "LEFT", "warna": "FF0000"},
        {"kunci": "Ttd. Komandan", "align": "RIGHT"},
    ],
    "tabel": 1, "gambar": 1, "teks_harus_ada": ["seluruh anggota"],
    "teks_harus_tidak_ada": ["\n3\n"],
})

# ---- 2. dua kolom ----
d = fitz.open()
p = d.new_page()
p.insert_text((200, 60), "BULETIN DUA KOLOM", fontsize=16)
kiri = ("Kolom kiri memuat berita utama tentang apel pagi yang diikuti seluruh personel. "
        "Kegiatan berlangsung tertib dan khidmat. Amanat pembina apel menekankan disiplin.")
kanan = ("Kolom kanan berisi agenda minggu ini yaitu latihan menembak pada hari Rabu dan rapat "
         "evaluasi pada hari Jumat. Seluruh perwira wajib hadir tepat waktu.")
tulis_terbungkus(p, kiri, 50, 110, 200)
tulis_terbungkus(p, kanan, 320, 110, 200)
simpan(d, "02-kolom.pdf", {
    "paragraf": [{"kunci": "BULETIN DUA KOLOM", "align": "HEADING"}],
    "tabel": 0, "gambar": 0,
    "urutan": ["apel pagi", "latihan menembak"],
})

# ---- 3. tanda hubung ----
d = fitz.open()
p = d.new_page()
p.insert_text((72, 100), "Hasil pemantau-", fontsize=11)
p.insert_text((72, 115), "an wilayah menunjukkan pela-", fontsize=11)
p.insert_text((72, 130), "ksanaan tugas berjalan baik.", fontsize=11)
simpan(d, "03-hubung.pdf", {
    "paragraf": [{"kunci": "Hasil pemantauan wilayah", "align": "LEFT"}],
    "tabel": 0, "gambar": 0,
    "teks_harus_ada": ["pemantauan", "pelaksanaan"],
    "teks_harus_tidak_ada": ["pemantau-", "pela-"],
})

# ---- 4. daftar ----
d = fitz.open()
p = d.new_page()
p.insert_text((72, 80), "DAFTAR HADIR", fontsize=15)
p.insert_text((72, 115), "• Apel pagi dilaksanakan tepat waktu", fontsize=11)
p.insert_text((72, 130), "• Seluruh personel wajib hadir", fontsize=11)
p.insert_text((72, 160), "1. Pembina apel memberi amanat", fontsize=11)
p.insert_text((72, 175), "2. Doa penutup oleh rohaniawan", fontsize=11)
simpan(d, "04-daftar.pdf", {
    "paragraf": [
        {"kunci": "DAFTAR HADIR", "align": "HEADING"},
        {"kunci": "Apel pagi", "align": "LIST"},
        {"kunci": "Pembina apel", "align": "LIST"},
    ],
    "tabel": 0, "gambar": 0,
})

# ---- 5. kop berulang 2 halaman ----
d = fitz.open()
for i in range(2):
    p = d.new_page()
    p.insert_text((72, 50), "MARKAS BESAR TNI - SATUAN SIBER", fontsize=12)
    tulis_terbungkus(p, f"Isi halaman {i + 1} tentang laporan kegiatan rutin satuan yang "
                        "dilaksanakan sesuai rencana kerja mingguan yang telah ditetapkan.", 72, 90, 450)
    p.insert_text((270, 800), str(i + 1), fontsize=10)
simpan(d, "05-kop.pdf", {
    "paragraf": [{"kunci": "laporan kegiatan rutin", "align": "LEFT"}],
    "tabel": 0, "gambar": 0,
    "header_harus_ada": ["MARKAS BESAR TNI"],
    "kop_tidak_di_badan": ["MARKAS BESAR TNI - SATUAN SIBER"],
})

# ---- 6. format campuran: tebal + miring + ukuran ----
d = fitz.open()
p = d.new_page()
p.insert_text((72, 100), "PENGUMUMAN PENTING", fontsize=14, fontname="hebo")
p.insert_text((72, 130), "Jadwal latihan menembak dimajukan.", fontsize=11, fontname="hebo")
p.insert_text((72, 160), "Peserta wajib membawa perlengkapan.", fontsize=11, fontname="heit")
simpan(d, "06-format.pdf", {
    "paragraf": [
        {"kunci": "PENGUMUMAN PENTING", "align": "HEADING"},
        {"kunci": "Jadwal latihan", "align": "LEFT", "bold": True},
        {"kunci": "wajib membawa", "align": "LEFT", "italic": True},
    ],
    "tabel": 0, "gambar": 0,
})

# ---- 7. justify asli: tepi kiri-kanan sejajar presisi ----
d = fitz.open()
p = d.new_page()
p.insert_text((150, 70), "SURAT EDARAN", fontsize=16)
t7 = ("seluruh personel diwajibkan mengikuti apel pagi setiap hari senin tepat pukul tujuh "
      "pagi di lapangan utama satuan dengan mengenakan seragam dinas lengkap dan atribut "
      "sesuai ketentuan yang berlaku tanpa terkecuali bagi semua golongan dan pangkat.")
X0, X1, FS, Y = 72, 522, 11, 110
kata, baris = t7.split(), []
lebar = lambda s: fitz.get_text_length(s, fontsize=FS)
for w in kata:
    t = " ".join(baris + [w])
    if lebar(t) > X1 - X0 and baris:
        n = len(baris)
        spasi = ((X1 - X0) - sum(lebar(k) for k in baris) - 2) / max(n - 1, 1)
        x = X0
        for k in baris:
            p.insert_text((x, Y), k, fontsize=FS)
            x += lebar(k) + lebar(" ") + spasi
        Y += FS + 3
        baris = [w]
    else:
        baris.append(w)
p.insert_text((X0, Y), " ".join(baris), fontsize=FS)
simpan(d, "07-justify.pdf", {
    "paragraf": [
        {"kunci": "SURAT EDARAN", "align": "HEADING"},
        {"kunci": "apel pagi", "align": "JUSTIFY"},
    ],
    "tabel": 0, "gambar": 0,
})

# ---- 8. hubung unicode (soft-hyphen jurnal ilmiah) ----
d = fitz.open()
p = d.new_page()
p.insert_text((72, 100), "Hasil transformasi organi­", fontsize=11)
p.insert_text((72, 115), "sasi menunjukkan implemen­", fontsize=11)
p.insert_text((72, 130), "tasi sistem berjalan baik.", fontsize=11)
simpan(d, "08-hubung-unicode.pdf", {
    "paragraf": [{"kunci": "Hasil transformasi organisasi", "align": "LEFT"}],
    "tabel": 0, "gambar": 0,
    "teks_harus_ada": ["transformasi", "organisasi", "implementasi"],
    "teks_harus_tidak_ada": ["organi­sasi", "implemen­tasi"],
})

# ---- 9. kop surat dinas (tengah + tanggal kanan + tembusan) ----
d = fitz.open()
p = d.new_page()
tulis_tengah(p, "MARKAS BESAR TENTARA NASIONAL INDONESIA", 60)
tulis_tengah(p, "SATUAN SIBER", 78)
p.insert_text((400, 130), "Jakarta, 5 Oktober 2026", fontsize=11)
tulis_terbungkus(p, "Bersama ini disampaikan laporan hasil kegiatan pemantauan wilayah "
                    "yang dilaksanakan oleh tim gabungan pada tanggal tersebut di atas "
                    "dengan hasil seluruh pos dalam keadaan aman dan terkendali.", 72, 170, 450)
p.insert_text((380, 300), "Tembusan :", fontsize=11)
p.insert_text((380, 316), "1. Panglima TNI", fontsize=11)
p.insert_text((380, 332), "2. Kasum TNI", fontsize=11)
simpan(d, "09-kopsurat.pdf", {
    "paragraf": [
        {"kunci": "MARKAS BESAR", "align": "CENTER"},
        {"kunci": "Jakarta, 5 Oktober 2026", "align": "RIGHT"},
        # AMBIGU by-design: tepi sejajar menyerupai justify; LEFT/JUSTIFY diterima
        {"kunci": "pemantauan wilayah", "align_any": ["LEFT", "JUSTIFY"]},
        {"kunci": "Panglima TNI", "align": "LIST"},
    ],
    "tabel": 0, "gambar": 0,
    "teks_harus_ada": ["Jakarta, 5 Oktober 2026", "Kasum TNI"],
})

# ---- 10. catatan kaki + referensi kecil ----
d = fitz.open()
p = d.new_page()
p.insert_text((200, 70), "LAPORAN SINGKAT", fontsize=16)
tulis_terbungkus(p, "Kegiatan evaluasi dilaksanakan sesuai rencana kerja yang telah "
                    "ditetapkan oleh pimpinan satuan pada awal tahun anggaran berjalan.",
                 72, 120, 450)
p.insert_text((72, 700), "1 Data diperoleh dari staf operasi dan logistik.", fontsize=8)
p.insert_text((72, 712), "2 Wawancara dilakukan pada bulan September 2026.", fontsize=8)
simpan(d, "10-catatan-kaki.pdf", {
    "paragraf": [
        {"kunci": "LAPORAN SINGKAT", "align": "HEADING"},
        {"kunci": "Kegiatan evaluasi", "align": "LEFT"},
    ],
    "tabel": 0, "gambar": 0,
    "teks_harus_ada": ["staf operasi", "September 2026"],
})

# ---- 11. jurnal mini 2 kolom + hubung + daftar pustaka ----
d = fitz.open()
p = d.new_page()
p.insert_text((230, 50), "Mini Journal", fontsize=15)
kj = ("Penelitian ini membahas transformasi digital dalam organisasi modern. "
      "Hasil menunjukkan implementasi sistem berjalan dengan baik dan efektif.")
tulis_terbungkus(p, kj, 50, 100, 200)
tulis_terbungkus(p, kj, 320, 100, 200)
p.insert_text((50, 400), "REFERENCES", fontsize=14, fontname="hebo")
p.insert_text((50, 420), "[1] Smith J. Digital systems. 2023.", fontsize=10)
p.insert_text((50, 434), "[2] Doe A. Transformation. 2024.", fontsize=10)
simpan(d, "11-jurnal-mini.pdf", {
    "paragraf": [
        {"kunci": "Mini Journal", "align": "HEADING"},
        {"kunci": "transformasi digital", "align": "LEFT"},
        {"kunci": "REFERENCES", "align": "HEADING"},
    ],
    "tabel": 0, "gambar": 0,
    "urutan": ["membahas transformasi", "Smith J"],
    "teks_harus_ada": ["Digital systems", "Transformation. 2024"],
})

# ---- 12. kop lebar + blok tanda tangan (gaya SKET) ----
d = fitz.open()
p = d.new_page()  # 612 x 1008
tulis_tengah(p, "MINISTRY OF DEFENSE OF THE REPUBLIC OF INDONESIA", 60)
tulis_tengah(p, "THE REPUBLIC OF INDONESIA DEFENSE UNIVERSITY", 80)
tulis_tengah(p, "OFFICIAL MEMO NUMBER 3526", 120)
tulis_terbungkus(p, "This letter is issued pursuant to the decree of the rector "
                    "concerning student status for the academic year 2025 slash 2026 "
                    "and is valid for all administrative purposes.", 72, 170, 450)
p.insert_text((380, 320), "Bogor, 22 February 2026", fontsize=11)
p.insert_text((360, 360), "Prof. Dr. Henry Setiyanto", fontsize=11)
simpan(d, "12-kop-lebar.pdf", {
    "paragraf": [
        {"kunci": "MINISTRY OF DEFENSE", "align": "CENTER"},
        {"kunci": "DEFENSE UNIVERSITY", "align": "CENTER"},
        {"kunci": "OFFICIAL MEMO NUMBER", "align": "CENTER"},
        {"kunci": "pursuant to the decree", "align": "LEFT"},
        {"kunci": "Bogor, 22 February 2026", "align": "RIGHT"},
        {"kunci": "Henry Setiyanto", "align": "RIGHT"},
    ],
    "tabel": 0, "gambar": 0,
})

# ---- 13. pangkat atas (super/subskrip sebaris) ----
d = fitz.open()
p = d.new_page()
_hv = fitz.Font("helv")
_tw = fitz.TextWriter(p.rect)
_tw.append((72, 100), "Catatan kaki", font=_hv, fontsize=11)
_tw.append((165, 96), "1", font=_hv, fontsize=7)
_tw.append((172, 100), " berlaku mulai Januari.", font=_hv, fontsize=11)
_tw.append((72, 130), "Rumus air H", font=_hv, fontsize=11)
_tw.append((165, 126), "2", font=_hv, fontsize=7)
_tw.append((172, 130), "O sangat penting.", font=_hv, fontsize=11)
_tw.write_text(p)
simpan(d, "13-pangkat.pdf", {
    "paragraf": [
        {"kunci": "Catatan kaki", "align": "LEFT", "super": "1"},
        {"kunci": "Rumus air H", "align": "LEFT", "super": "2"},
    ],
    "tabel": 0, "gambar": 0,
    "teks_harus_ada": ["berlaku mulai Januari", "sangat penting"],
})

# ---- 14. tabel bergaris + teks sekitar (anti duplikasi) ----
d = fitz.open()
p = d.new_page()
p.insert_text((72, 80), "Pengantar singkat sebelum tabel.", fontsize=11)
for r, row in enumerate([("Nama", "Unit", "Ket"), ("Budi", "Siber", "Hadir"), ("Sari", "Umum", "Izin")]):
    for c, v in enumerate(row):
        rc = fitz.Rect(72 + c * 140, 120 + r * 22, 72 + (c + 1) * 140, 120 + (r + 1) * 22)
        p.draw_rect(rc)
        p.insert_text((rc.x0 + 4, rc.y0 + 15), v, fontsize=11)
p.insert_text((72, 230), "Penutup sesudah tabel.", fontsize=11)
simpan(d, "14-tabel-teks.pdf", {
    "paragraf": [
        {"kunci": "Pengantar singkat", "align": "LEFT"},
        {"kunci": "Penutup sesudah", "align": "LEFT"},
    ],
    "tabel": 1, "gambar": 0,
    "teks_harus_ada": ["Pengantar singkat", "Penutup sesudah"],
    "teks_tidak_ganda": ["Budi", "Siber", "Sari"],
})

# ---- 15. watermark raksasa-pudar + stempel miring ----
d = fitz.open()
p = d.new_page()
p.insert_text((100, 220), "RAHASIA", fontsize=60, color=(0.82, 0.82, 0.82))
tulis_terbungkus(p, "Isi dokumen penting yang wajib dibaca oleh seluruh personel "
                    "dan dilaksanakan dengan penuh tanggung jawab.", 72, 120, 450)
tulis_terbungkus(p, "Demikian disampaikan untuk menjadi perhatian bersama.", 72, 220, 450)
_sh = p.new_shape()
_sh.insert_textbox(fitz.Rect(480, 600, 540, 780), "DITERIMA", fontsize=14,
                   fontname="helv", rotate=90)
_sh.commit()
simpan(d, "15-watermark.pdf", {
    "paragraf": [
        {"kunci": "wajib dibaca", "align": "LEFT"},
    ],
    "tabel": 0, "gambar": 0,
    "tidak_di_badan": ["RAHASIA", "DITERIMA"],
})

# ---- 16. CV + watermark diagonal + foto + kolom ganda (replika acuan user) ----
d = fitz.open()
p = d.new_page(width=595, height=842)
tulis_tengah(p, "Rizky Ramadhan Nasution", 60, size=20)
p.insert_text((150, 92), "ramadanrizky941@gmail.com", fontsize=10)
tulis_terbungkus(p, "Lulusan Informatika Universitas Pertahanan dengan minat analisis "
                    "data dan pengembangan sistem.", 72, 130, 430)
p.insert_text((72, 200), "TECHNICAL SKILLS", fontsize=13)
tulis_terbungkus(p, "Public Speaking Data Science Programming", 72, 225, 200)
tulis_terbungkus(p, "Game Engineering Copywriting Cloud Systems", 320, 225, 200)
p.insert_text((72, 330), "ACHIEVEMENTS", fontsize=13)
p.insert_text((72, 355), "Gold Medal National Olympiad", fontsize=11)
p.insert_text((430, 355), "2022", fontsize=11)
p.insert_text((72, 375), "Second Place Speech Contest", fontsize=11)
p.insert_text((430, 375), "2022", fontsize=11)
# foto kanan atas (raster merah)
import fitz as _fz
_foto = _fz.Pixmap(_fz.csRGB, fitz.Rect(0, 0, 90, 120), 0)
_foto.set_rect(_foto.irect, (180, 30, 60))
p.insert_image(fitz.Rect(440, 30, 530, 150), pixmap=_foto)
# watermark diagonal RAHASIA pudar menindih isi (sudut -45)
_sh = p.new_shape()
_m = fitz.Matrix(0.7071, 0.7071, -0.7071, 0.7071, 0, 0)
_sh.insert_text(fitz.Point(120, 550), "RAHASIA", fontsize=64, fontname="helv",
                color=(0.80, 0.80, 0.80), morph=(fitz.Point(120, 550), _m))
_sh.commit()
simpan(d, "16-cv-watermark.pdf", {
    "paragraf": [
        {"kunci": "Rizky Ramadhan", "align": "HEADING"},
        {"kunci": "Gold Medal", "align": "LEFT"},
    ],
    "tabel": 0, "gambar": 1,
    "teks_harus_ada": ["TECHNICAL SKILLS", "ACHIEVEMENTS", "2022"],
    "tidak_di_badan": ["RAHASIA"],
})

with open(os.path.join(KORPUS, 'labels.json'), 'w', encoding='utf-8') as fh:
    json.dump(LABELS, fh, ensure_ascii=False, indent=2)
print("labels.json OK:", len(LABELS["kasus"]), "kasus")

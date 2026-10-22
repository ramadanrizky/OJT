"""Korpus latih OCR: render teks jadi gambar (simulasi scan) + ground truth.
Jalankan: python pelatihan_konversi/buat_korpus_ocr.py
Hasil: pelatihan_konversi/korpus_ocr/*.pdf + labels_ocr.json
"""
import json
import os

import fitz

DASAR = os.path.dirname(os.path.abspath(__file__))
KORPUS = os.path.join(DASAR, 'korpus_ocr')
os.makedirs(KORPUS, exist_ok=True)

TEKS = [
    ("memo dinas", "MEMORANDUM DINAS",
     ["Kepala Bagian Operasi kepada seluruh personel.",
      "Bersama ini disampaikan jadwal apel pagi pada hari Senin.",
      "Apel dilaksanakan pukul tujuh tepat di lapangan utama.",
      "Seluruh personel wajib hadir dengan seragam dinas."]),
    ("surat undangan", "SURAT UNDANGAN RAPAT",
     ["Nomor B/12/VIII/2026 tanggal 3 Oktober 2026.",
      "Rapat evaluasi mingguan dilaksanakan hari Jumat.",
      "Tempat di ruang rapat utama lantai dua.",
      "Peserta seluruh perwira dan bintara tinggi."]),
    ("pengumuman", "PENGUMUMAN HASIL SELEKSI",
     ["Panitia mengumumkan hasil seleksi administrasi.",
      "Peserta yang lulus agar hadir pada tahap berikutnya.",
      "Membawa kartu identitas dan berkas asli.",
      "Jadwal menyusul melalui pengumuman resmi."]),
]

LABELS = {"kasus": []}

for nama, judul, paras in TEKS:
    # 1) tulis docx-like via PDF teks biasa dulu
    tmp = fitz.open()
    pg = tmp.new_page(width=595, height=842)
    pg.insert_text((150, 80), judul, fontsize=16)
    y = 130
    for pr in paras:
        pg.insert_text((72, y), pr, fontsize=12)
        y += 24
    # 2) rasterisasi jadi gambar (simulasi hasil scan 150dpi)
    pix = pg.get_pixmap(dpi=150)
    d2 = fitz.open()
    p2 = d2.new_page(width=595, height=842)
    p2.insert_image(p2.rect, pixmap=pix)
    fn = f"scan-{nama.replace(' ', '-')}.pdf"
    d2.save(os.path.join(KORPUS, fn))
    gt = judul + "\n" + "\n".join(paras)
    LABELS["kasus"].append({"pdf": fn, "ground_truth": gt})
    print("korpus:", fn)

with open(os.path.join(KORPUS, 'labels_ocr.json'), 'w', encoding='utf-8') as fh:
    json.dump(LABELS, fh, ensure_ascii=False, indent=2)
print("labels_ocr.json OK")

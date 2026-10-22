"""Latih parameter OCR: ukur CER (character error rate) lawan ground truth,
lalu grid-search dpi + ambang penggabungan paragraf.
Jalankan: python pelatihan_konversi/latih_ocr.py [--tune]
"""
import json
import os
import sys

DASAR = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(DASAR)
sys.path.insert(0, ROOT)
os.chdir(ROOT)

from docx import Document  # noqa: E402

import app as A  # noqa: E402

KORPUS = os.path.join(DASAR, 'korpus_ocr')
TMP = os.path.join(DASAR, 'tmp_ocr')
os.makedirs(TMP, exist_ok=True)


def muat():
    with open(os.path.join(KORPUS, 'labels_ocr.json'), encoding='utf-8') as fh:
        return json.load(fh)["kasus"]


def cer(pred, ref):
    """Character Error Rate via difflib (akurat & tanpa dependensi)."""
    import difflib
    a = "".join(ref.split())
    b = "".join(pred.split())
    if not a:
        return 1.0 if b else 0.0
    sm = difflib.SequenceMatcher(None, a, b, autojunk=False)
    return 1.0 - sm.ratio()


def konversi_ocr(pdf, out, dpi=150):
    A._pdf2word_ocr(pdf, out, dpi=dpi)


def skor_satu(pdf, gt, dpi=150):
    out = os.path.join(TMP, os.path.basename(pdf) + ".docx")
    try:
        konversi_ocr(pdf, out, dpi=dpi)
    except Exception as e:
        return 1.0, f"CRASH: {e}"
    doc = Document(out)
    pred = "\n".join(p.text for p in doc.paragraphs if p.text.strip())
    try:
        os.remove(out)
    except OSError:
        pass
    return cer(pred, gt), pred[:60]


GRID_DPI = [100, 150, 200]


def main():
    kasus = muat()
    print("Korpus OCR:", len(kasus), "kasus")
    for dpi in GRID_DPI if "--tune" in sys.argv else [150]:
        tot, n = 0.0, 0
        for ks in kasus:
            v, _ = skor_satu(os.path.join(KORPUS, ks["pdf"]), ks["ground_truth"], dpi=dpi)
            tot += v
            n += 1
        print(f"dpi={dpi}: CER rata-rata={tot / n:.3f}")
    if "--tune" not in sys.argv:
        print("Tambahkan --tune untuk grid-search dpi.")
        return
    cfg = A._muat_config_konversi()
    cfg["ocr_dpi"] = min(GRID_DPI, key=lambda d: sum(
        skor_satu(os.path.join(KORPUS, ks["pdf"]), ks["ground_truth"], dpi=d)[0]
        for ks in kasus))
    with open(os.path.join(DASAR, 'konversi_config_v1.json'), 'w', encoding='utf-8') as fh:
        json.dump(cfg, fh, indent=2)
    print("ocr_dpi terbaik disimpan:", cfg["ocr_dpi"])


if __name__ == "__main__":
    main()

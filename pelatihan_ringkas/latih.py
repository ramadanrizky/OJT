"""Melatih AI ringkas: ukur baseline vs model baru (ROUGE), lalu tuning bobot.

Jalankan:  py pelatihan_ringkas\\latih.py            (evaluasi saja)
           py pelatihan_ringkas\\latih.py --tune      (grid-search + simpan config)
"""
import itertools
import json
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
os.chdir(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import ringkas_ai
from pelatihan_ringkas.ringkas_ai import (
    DEFAULT_CONFIG, bersihkan_teks, split_kalimat_id, skor_kalimat, MODES,
)

KORPUS = os.path.join('pelatihan_ringkas', 'korpus')


def token(s):
    return re.findall(r'[a-z]+', s.lower())


def rouge_n_f1(calon, acuan, n=1):
    from collections import Counter
    ct, at = token(calon), token(acuan)
    if not ct or not at:
        return 0.0
    cg = Counter(tuple(ct[i:i + n]) for i in range(len(ct) - n + 1))
    ag = Counter(tuple(at[i:i + n]) for i in range(len(at) - n + 1))
    overlap = sum(min(cg[g], ag[g]) for g in cg)
    if not overlap:
        return 0.0
    p = overlap / max(sum(cg.values()), 1)
    r = overlap / max(sum(ag.values()), 1)
    return 2 * p * r / (p + r)


def lcs(a, b):
    m, n = len(a), len(b)
    if not m or not n:
        return 0
    prev = [0] * (n + 1)
    for i in range(1, m + 1):
        cur = [0] * (n + 1)
        ai = a[i - 1]
        for j in range(1, n + 1):
            cur[j] = prev[j - 1] + 1 if ai == b[j - 1] else max(prev[j], cur[j - 1])
        prev = cur
    return prev[n]


def rouge_l_f1(calon, acuan):
    ct, at = token(calon), token(acuan)
    if not ct or not at:
        return 0.0
    l = lcs(ct, at)
    p, r = l / len(ct), l / len(at)
    return 0.0 if p + r == 0 else 2 * p * r / (p + r)


def muat_korpus():
    with open(os.path.join(KORPUS, 'references.json'), encoding='utf-8') as fh:
        refs = json.load(fh)
    data = []
    for fn, ref in refs.items():
        with open(os.path.join(KORPUS, fn), encoding='utf-8') as fh:
            data.append((fn, fh.read(), ref))
    # tambah dokumen nyata dari uploads (tanpa referensi manual -> lewati skor, cek fragmentasi saja)
    return data


def ringkas_dengan_cfg(teks, cfg, mode='sedang', _cache=None):
    if _cache is None:
        bersih = bersihkan_teks(teks)
        kalimat = split_kalimat_id(bersih)
    else:
        bersih, kalimat = _cache[0], _cache[1]
    total = len(kalimat)
    if total <= 2:
        return bersih, kalimat
    m = MODES[mode]
    jumlah = max(2, min(m['maks'], round(total * m['rasio'])))
    if _cache is not None and len(_cache) == 3:
        skor = skor_kalimat(kalimat, cfg, token_docs=_cache[2])
    else:
        skor = skor_kalimat(kalimat, cfg)
    pilih = sorted(sorted(range(total), key=lambda i: skor[i], reverse=True)[:jumlah])
    return '\n\n'.join(kalimat[i] for i in pilih), kalimat


def skor_model(cfg, data, mode='sedang', caches=None):
    r1 = rl = frag = 0.0
    for idx, (_, teks, ref) in enumerate(data):
        hasil, kalimat = ringkas_dengan_cfg(teks, cfg, mode,
                                            _cache=caches[idx] if caches else None)
        r1 += rouge_n_f1(hasil, ref)
        rl += rouge_l_f1(hasil, ref)
        # fragmentasi: proporsi kalimat hasil yang < 4 kata ("sepotong kata")
        hs = [k for k in hasil.split('\n\n') if k.strip()]
        frag += sum(1 for k in hs if len(k.split()) < 4) / max(len(hs), 1)
    n = len(data)
    return r1 / n, rl / n, frag / n


def baseline_sumy(teks):
    """Baseline lama: LexRank sumy + tokenizer Inggris (cara kode saat ini)."""
    from sumy.parsers.plaintext import PlaintextParser
    from sumy.nlp.tokenizers import Tokenizer
    from sumy.summarizers.lex_rank import LexRankSummarizer
    from Sastrawi.Stemmer.StemmerFactory import StemmerFactory
    from Sastrawi.StopWordRemover.StopWordRemoverFactory import StopWordRemoverFactory
    parser = PlaintextParser.from_string(teks, Tokenizer('english'))
    sents = list(parser.document.sentences)
    if len(sents) <= 2:
        return teks
    summ = LexRankSummarizer(StemmerFactory().create_stemmer().stem)
    summ.stop_words = StopWordRemoverFactory().get_stop_words()
    jumlah = max(2, min(10, round(len(sents) * 0.35)))
    return '\n\n'.join(str(s) for s in summ(parser.document, jumlah))


def main():
    data = muat_korpus()
    print(f'Korpus: {len(data)} dokumen')
    b1 = bl = bf = 0.0
    for _, teks, ref in data:
        h = baseline_sumy(teks)
        b1 += rouge_n_f1(h, ref)
        bl += rouge_l_f1(h, ref)
        hs = [k for k in h.split('\n\n') if k.strip()]
        bf += sum(1 for k in hs if len(k.split()) < 4) / max(len(hs), 1)
    n = len(data)
    print(f'BASELINE sumy : ROUGE-1={b1/n:.3f} ROUGE-L={bl/n:.3f} fragmentasi={bf/n:.3f}')

    cfg = ringkas_ai.muat_config()
    r1, rl, fg = skor_model(cfg, data)
    print(f'MODEL baru    : ROUGE-1={r1:.3f} ROUGE-L={rl:.3f} fragmentasi={fg:.3f}')

    if '--tune' not in sys.argv:
        print('Tambahkan --tune untuk grid-search bobot.')
        return
    # Pra-hitung sekali (stemming mahal): bersih + kalimat + token
    from pelatihan_ringkas.ringkas_ai import _sastrawi, _tokenize
    stemmer, stopwords = _sastrawi()
    caches = []
    for _, teks, _ in data:
        b = bersihkan_teks(teks)
        k = split_kalimat_id(b)
        caches.append((b, k, [_tokenize(x, stemmer, stopwords) for x in k]))
    terbaik, skor_terbaik = None, -1.0
    grid = {
        'bobot_pagerank': [0.40, 0.55, 0.70],
        'bobot_posisi': [0.05, 0.15, 0.25],
        'bobot_panjang': [0.05, 0.10, 0.20],
        'bobot_kata_kunci': [0.05, 0.12, 0.20],
        'bobot_angka': [0.0, 0.08, 0.15],
        'ambang_kemiripan': [0.05, 0.10, 0.20],
    }
    kunci = list(grid)
    total = 1
    for k in kunci:
        total *= len(grid[k])
    print(f'Grid search {total} kombinasi...')
    i = 0
    for vals in itertools.product(*[grid[k] for k in kunci]):
        i += 1
        c = dict(DEFAULT_CONFIG)
        c.update(dict(zip(kunci, vals)))
        # normalisasi bobot skor agar jumlah = 1
        s = c['bobot_pagerank'] + c['bobot_posisi'] + c['bobot_panjang'] + c['bobot_kata_kunci'] + c['bobot_angka']
        for k in ('bobot_pagerank', 'bobot_posisi', 'bobot_panjang', 'bobot_kata_kunci', 'bobot_angka'):
            c[k] = round(c[k] / s, 3)
        r1, rl, fg = skor_model(c, data, caches=caches)
        gab = (r1 + rl) / 2 - fg * 0.3  # penalti fragmentasi
        if gab > skor_terbaik:
            skor_terbaik, terbaik = gab, dict(c)
            print(f'  [{i}/{total}] BARU TERBAIK gab={gab:.3f} (R1={r1:.3f} RL={rl:.3f} frag={fg:.3f})')
    with open(os.path.join('pelatihan_ringkas', 'ringkas_config_v1.json'), 'w', encoding='utf-8') as fh:
        json.dump(terbaik, fh, indent=2)
    print('Config terbaik disimpan:', terbaik)


if __name__ == '__main__':
    main()

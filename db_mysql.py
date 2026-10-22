# -*- coding: utf-8 -*-
"""
Adapter MySQL/MariaDB dengan API mirip Cloud Firestore.

Tujuan: mengganti database Firestore ke MySQL TANPA mengubah kode pemakaian
di `app.py` (129 titik `fdb.collection(...)`) dan tanpa mengubah bentuk data
yang dikirim ke frontend.

Konsep penyimpanan:
  - 1 dokumen Firestore  ==  1 baris MySQL
  - Kolom `id`   : ID dokumen Firestore (dipertahankan apa adanya)
  - Kolom `data` : seluruh isi dokumen dalam bentuk JSON (LONGTEXT)
  - Kolom lain   : salinan field yang sering dipakai `.where()`, hanya untuk
                   indeks/pencarian. Sumber kebenaran tetap kolom `data`.

Dengan pola ini, field apa pun yang ada di Firestore tetap utuh, termasuk
field yang belum terpetakan, array, objek bersarang, dan datetime.

Konfigurasi (environment variable), default cocok untuk XAMPP:
    MYSQL_HOST      default 127.0.0.1
    MYSQL_PORT      default 3306
    MYSQL_USER      default root
    MYSQL_PASSWORD  default (kosong)
    MYSQL_DB        default portal_karyawan

Dependensi: pip install PyMySQL
"""

from __future__ import annotations

import datetime as _dt
import decimal as _decimal
import functools as _functools
import json as _json
import os as _os
import re as _re
import secrets as _secrets
import string as _string
import threading as _threading

import pymysql
from pymysql.cursors import DictCursor as _DictCursor


# =============================================================================
# KONFIGURASI KONEKSI
# =============================================================================
KONFIG = {
    "host": _os.environ.get("MYSQL_HOST", "127.0.0.1"),
    "port": int(_os.environ.get("MYSQL_PORT", "3306")),
    "user": _os.environ.get("MYSQL_USER", "root"),
    "password": _os.environ.get("MYSQL_PASSWORD", ""),
    "database": _os.environ.get("MYSQL_DB", "portal_karyawan"),
    "charset": "utf8mb4",
}


# =============================================================================
# SKEMA TABEL
# Nama tabel = nama koleksi Firestore.
# Kolom tambahan = field yang dipakai pada `.where(...)` agar bisa diindeks.
# =============================================================================
TABEL = {
    "users": {
        "email": "VARCHAR(190)",
        "role": "VARCHAR(20)",
        "superior_id": "VARCHAR(191)",
        "fullname": "VARCHAR(190)",
        "telegram_chat_id": "VARCHAR(64)",
    },
    "struktur": {
        "user_id": "VARCHAR(191)",
        "parent_id": "VARCHAR(191)",
        "nama": "VARCHAR(190)",
        "urutan": "INT",
    },
    "tugas": {
        "user_id": "VARCHAR(191)",
        "assigned_by": "VARCHAR(191)",
        "status": "VARCHAR(20)",
    },
    "disposisi_surat": {
        "atasan_id": "VARCHAR(191)",
        "pengirim_id": "VARCHAR(191)",
        "asal_surat_id": "VARCHAR(191)",
        "status": "VARCHAR(30)",
        "file_nama_simpan": "VARCHAR(255)",
        "lampiran_nama_simpan": "VARCHAR(255)",
    },
    "disposisi_log": {
        "surat_id": "VARCHAR(191)",
    },
    "disposisi_hist": {},
    "bantuan_tiket": {
        "user_id": "VARCHAR(191)",
        "status": "VARCHAR(30)",
    },
    "pengumuman": {},
    "jadwal_rapat": {},
    "format_surat": {},
    "password_reset_otp": {
        "email": "VARCHAR(190)",
    },
}

# (nama_index, kolom, unik?)
INDEKS = {
    "users": [("idx_users_email", "email", False),
              ("idx_users_role", "role", False),
              ("idx_users_superior", "superior_id", False)],
    "struktur": [("idx_struktur_user", "user_id", False),
                 ("idx_struktur_parent", "parent_id", False)],
    "tugas": [("idx_tugas_user", "user_id", False),
              ("idx_tugas_assigned", "assigned_by", False),
              ("idx_tugas_status", "status", False)],
    "disposisi_surat": [("idx_disp_atasan", "atasan_id", False),
                        ("idx_disp_pengirim", "pengirim_id", False),
                        ("idx_disp_asal", "asal_surat_id", False),
                        ("idx_disp_file", "file_nama_simpan", False),
                        ("idx_disp_lampiran", "lampiran_nama_simpan", False)],
    "disposisi_log": [("idx_displog_surat", "surat_id", False)],
    "bantuan_tiket": [("idx_tiket_user", "user_id", False),
                      ("idx_tiket_status", "status", False)],
    "password_reset_otp": [("idx_otp_email", "email", False)],
}

_VALID_KOLOM = _re.compile(r"^[A-Za-z0-9_]+$")


# =============================================================================
# SENTINEL (pengganti firestore.SERVER_TIMESTAMP / DELETE_FIELD / ArrayUnion)
# =============================================================================
class _Sentinel:
    def __init__(self, nama):
        self._nama = nama

    def __repr__(self):
        return f"<{self._nama}>"


SERVER_TIMESTAMP = _Sentinel("SERVER_TIMESTAMP")
DELETE_FIELD = _Sentinel("DELETE_FIELD")


class ArrayUnion:
    """Tambah nilai ke array tanpa duplikat (mirip firestore.ArrayUnion)."""

    def __init__(self, values):
        self.values = list(values or [])


class ArrayRemove:
    """Hapus nilai dari array (mirip firestore.ArrayRemove)."""

    def __init__(self, values):
        self.values = list(values or [])


# =============================================================================
# ENCODE / DECODE JSON
# Datetime disimpan dengan penanda agar saat dibaca kembali menjadi datetime
# (dibutuhkan fitur OTP yang membandingkan expires_at).
# =============================================================================
_TS_KEY = "__datetime__"
_TS_KEY_LAMA = "__tanggal__"


def _encode(nilai):
    if isinstance(nilai, _Sentinel):
        # SERVER_TIMESTAMP / DELETE_FIELD yang lolos ke data diubah jadi waktu.
        return {_TS_KEY: _dt.datetime.now().isoformat(sep=" ")}
    if isinstance(nilai, _dt.datetime):
        return {_TS_KEY: nilai.isoformat(sep=" ")}
    if isinstance(nilai, _dt.date):
        return {_TS_KEY_LAMA: nilai.isoformat()}
    if isinstance(nilai, _decimal.Decimal):
        return float(nilai)
    if isinstance(nilai, (bytes, bytearray)):
        return {"__bytes__": bytes(nilai).hex()}
    if isinstance(nilai, dict):
        return {str(k): _encode(v) for k, v in nilai.items()}
    if isinstance(nilai, (list, tuple, set)):
        return [_encode(v) for v in nilai]
    return nilai


def _decode(nilai):
    if isinstance(nilai, dict):
        if len(nilai) == 1:
            if _TS_KEY in nilai:
                try:
                    return _dt.datetime.fromisoformat(str(nilai[_TS_KEY]))
                except ValueError:
                    return str(nilai[_TS_KEY])
            if _TS_KEY_LAMA in nilai:
                try:
                    return _dt.date.fromisoformat(str(nilai[_TS_KEY_LAMA]))
                except ValueError:
                    return str(nilai[_TS_KEY_LAMA])
            if "__bytes__" in nilai:
                try:
                    return bytes.fromhex(str(nilai["__bytes__"]))
                except ValueError:
                    return nilai
        return {k: _decode(v) for k, v in nilai.items()}
    if isinstance(nilai, list):
        return [_decode(v) for v in nilai]
    return nilai


def _dump_json(data):
    return _json.dumps(_encode(data), ensure_ascii=False, separators=(",", ":"))


def _load_json(teks, default=None):
    if teks is None:
        return default
    if isinstance(teks, (bytes, bytearray)):
        teks = teks.decode("utf-8")
    try:
        return _decode(_json.loads(teks))
    except (ValueError, TypeError):
        return default


# =============================================================================
# KONEKSI (satu koneksi per thread, dipakai ulang)
# =============================================================================
_local = _threading.local()


def _koneksi_baru(pakai_database=True):
    konfig = dict(KONFIG)
    if not pakai_database:
        konfig.pop("database", None)
    return pymysql.connect(
        autocommit=True,
        cursorclass=_DictCursor,
        **konfig,
    )


def _koneksi():
    conn = getattr(_local, "conn", None)
    if conn is None:
        conn = _koneksi_baru()
        _local.conn = conn
    else:
        try:
            conn.ping(reconnect=True)
        except Exception:
            try:
                conn.close()
            except Exception:
                pass
            conn = _koneksi_baru()
            _local.conn = conn
    return conn


def tutup_koneksi_thread():
    """Tutup koneksi milik thread ini (dipakai saat shutdown/uji)."""
    conn = getattr(_local, "conn", None)
    if conn is not None:
        try:
            conn.close()
        finally:
            _local.conn = None


def _ddl(nama_tabel):
    kolom = TABEL[nama_tabel]
    baris = [
        "  `id` VARCHAR(191) NOT NULL",
        "  `data` LONGTEXT NOT NULL",
    ]
    for nama, tipe in kolom.items():
        baris.append(f"  `{nama}` {tipe} NULL")
    baris.append("  PRIMARY KEY (`id`)")
    for nama_idx, kol, unik in INDEKS.get(nama_tabel, []):
        awalan = "UNIQUE " if unik else ""
        baris.append(f"  {awalan}KEY `{nama_idx}` (`{kol}`)")
    isi = ",\n".join(baris)
    return (
        f"CREATE TABLE IF NOT EXISTS `{nama_tabel}` (\n{isi}\n) "
        "ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci"
    )


def _tabel_ada(cur, nama):
    cur.execute(
        "SELECT 1 FROM information_schema.TABLES "
        "WHERE TABLE_SCHEMA=%s AND TABLE_NAME=%s",
        (KONFIG["database"], nama),
    )
    return cur.fetchone() is not None


def _info_kolom(cur, nama):
    cur.execute(f"SHOW COLUMNS FROM `{nama}`")
    return {b["Field"]: b for b in cur.fetchall()}


def _nama_tabel_lama(cur, nama):
    kandidat = f"{nama}_lama"
    nomor = 1
    while _tabel_ada(cur, kandidat):
        nomor += 1
        kandidat = f"{nama}_lama_{nomor}"
    return kandidat


def siapkan_skema(verbose=True):
    """Buat database + tabel + kolom yang belum ada (aman dijalankan berulang).

    Tabel lama dari versi MySQL sebelumnya (mis. `tugas` dengan id INT dan
    tanpa kolom `data`) otomatis diubah namanya menjadi `<tabel>_lama` supaya
    datanya tidak hilang, lalu tabel baru dibuat. Ini mencegah error
    "Unknown column 'data' in 'field list'".
    """
    db = KONFIG["database"]
    if not _VALID_KOLOM.match(db):
        raise ValueError(f"Nama database tidak valid: {db!r}")

    conn = _koneksi_baru(pakai_database=False)
    try:
        with conn.cursor() as cur:
            cur.execute(
                f"CREATE DATABASE IF NOT EXISTS `{db}` "
                "CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci"
            )
    finally:
        conn.close()

    conn = _koneksi()
    with conn.cursor() as cur:
        for nama_tabel, kolom in TABEL.items():
            if _tabel_ada(cur, nama_tabel):
                info = _info_kolom(cur, nama_tabel)
                id_tipe = (info.get("id") or {}).get("Type", "").lower()
                if "data" not in info or not id_tipe.startswith(("varchar", "char")):
                    lama = _nama_tabel_lama(cur, nama_tabel)
                    cur.execute(f"RENAME TABLE `{nama_tabel}` TO `{lama}`")
                    if verbose:
                        print(
                            f"[INFO] Tabel lama `{nama_tabel}` (skema versi lama) "
                            f"diubah menjadi `{lama}`; membuat tabel baru."
                        )
                    cur.execute(_ddl(nama_tabel))
            else:
                cur.execute(_ddl(nama_tabel))

            # Tabel mungkin sudah ada dari versi skema sebelumnya: tambah kolom baru.
            ada = set(_info_kolom(cur, nama_tabel))
            for nama, tipe in kolom.items():
                if nama not in ada:
                    cur.execute(
                        f"ALTER TABLE `{nama_tabel}` ADD COLUMN `{nama}` {tipe} NULL"
                    )
            for nama_idx, kol, _unik in INDEKS.get(nama_tabel, []):
                cur.execute(f"SHOW INDEX FROM `{nama_tabel}` WHERE Key_name=%s", (nama_idx,))
                if not cur.fetchall():
                    cur.execute(
                        f"ALTER TABLE `{nama_tabel}` ADD INDEX `{nama_idx}` (`{kol}`)"
                    )
    if verbose:
        print(
            f"[OK] MySQL siap: {KONFIG['user']}@{KONFIG['host']}:"
            f"{KONFIG['port']}/{KONFIG['database']} ({len(TABEL)} tabel)"
        )


# =============================================================================
# QUERY
# =============================================================================
def _susun_where(nama_tabel, filter_list):
    bagian, params = [], []
    for field, op, nilai in filter_list:
        if not _VALID_KOLOM.match(str(field)):
            raise ValueError(f"Nama field tidak valid: {field!r}")
        kolom = field if field in TABEL[nama_tabel] else None
        if op == "==":
            if nilai is None:
                if kolom:
                    bagian.append(f"`{kolom}` IS NULL")
                else:
                    bagian.append(f"JSON_EXTRACT(`data`, %s) IS NULL")
                    params.append(f"$.{field}")
            elif kolom:
                bagian.append(f"`{kolom}` = %s")
                params.append(nilai)
            else:
                bagian.append("JSON_UNQUOTE(JSON_EXTRACT(`data`, %s)) = %s")
                params.extend([f"$.{field}", str(nilai)])
        elif op in ("<", "<=", ">", ">="):
            bagian.append(f"JSON_EXTRACT(`data`, %s) {op} %s")
            params.extend([f"$.{field}", nilai])
        elif op == "in":
            daftar = list(nilai or [])
            if not daftar:
                bagian.append("1=0")
                continue
            if kolom:
                tanda = ", ".join(["%s"] * len(daftar))
                bagian.append(f"`{kolom}` IN ({tanda})")
                params.extend(daftar)
            else:
                tanda = ", ".join(["%s"] * len(daftar))
                bagian.append(
                    f"JSON_UNQUOTE(JSON_EXTRACT(`data`, %s)) IN ({tanda})"
                )
                params.append(f"$.{field}")
                params.extend([str(v) for v in daftar])
        else:
            raise ValueError(f"Operator query tidak didukung: {op!r}")
    if not bagian:
        return "", []
    return " WHERE " + " AND ".join(bagian), params


# =============================================================================
# OBJEK-OBJEK API (meniru Firestore)
# =============================================================================
class DocumentSnapshot:
    def __init__(self, reference, data):
        self.reference = reference
        self._data = data

    @property
    def id(self):
        return self.reference.id

    @property
    def exists(self):
        return self._data is not None

    def to_dict(self):
        if self._data is None:
            return None
        # Salinan dangkal, sama seperti Firestore: pemanggil tidak bisa
        # mengubah data internal snapshot.
        return dict(self._data)

    def get(self, field):
        return (self._data or {}).get(field)


class DocumentReference:
    def __init__(self, tabel, doc_id):
        self._tabel = tabel
        self._id = doc_id

    @property
    def id(self):
        return self._id

    def get(self, transaction=None):
        conn = transaction._conn if transaction is not None else _koneksi()
        kunci = " FOR UPDATE" if transaction is not None else ""
        with conn.cursor() as cur:
            cur.execute(
                f"SELECT `data` FROM `{self._tabel}` WHERE `id`=%s{kunci}",
                (self._id,),
            )
            baris = cur.fetchone()
        data = _load_json(baris["data"]) if baris else None
        return DocumentSnapshot(self, data)

    def set(self, data, merge=False):
        conn = _koneksi()
        isi = dict(data or {})
        if merge:
            lama = self.get().to_dict() or {}
            lama.update(isi)
            isi = lama
        self._tulis(isi, conn)
        return self

    def _tulis(self, data, conn):
        baris = {"id": self._id, "data": _dump_json(data)}
        for kolom in TABEL[self._tabel]:
            nilai = data.get(kolom)
            if isinstance(nilai, _Sentinel):
                nilai = None
            elif isinstance(nilai, (dict, list, tuple)):
                nilai = _json.dumps(_encode(nilai), ensure_ascii=False)
            elif isinstance(nilai, _dt.datetime):
                nilai = nilai.strftime("%Y-%m-%d %H:%M:%S")
            elif isinstance(nilai, _dt.date):
                nilai = nilai.isoformat()
            baris[kolom] = nilai

        kolom_sql = ", ".join(f"`{k}`" for k in baris)
        placeholder = ", ".join(["%s"] * len(baris))
        pembaruan = ", ".join(
            f"`{k}`=VALUES(`{k}`)" for k in baris if k != "id"
        )
        sql = (
            f"INSERT INTO `{self._tabel}` ({kolom_sql}) VALUES ({placeholder}) "
            f"ON DUPLICATE KEY UPDATE {pembaruan}"
        )
        with conn.cursor() as cur:
            cur.execute(sql, list(baris.values()))

    def update(self, data):
        conn = _koneksi()
        luar = conn.get_autocommit()
        if luar:
            conn.begin()
        try:
            with conn.cursor() as cur:
                cur.execute(
                    f"SELECT `data` FROM `{self._tabel}` "
                    "WHERE `id`=%s FOR UPDATE",
                    (self._id,),
                )
                baris = cur.fetchone()
            lama = _load_json(baris["data"]) if baris else {}
            if lama is None:
                lama = {}
            baru = dict(lama)
            for kunci, nilai in (data or {}).items():
                if nilai is DELETE_FIELD:
                    baru.pop(kunci, None)
                elif isinstance(nilai, ArrayUnion):
                    sekarang = list(baru.get(kunci) or [])
                    for item in nilai.values:
                        if item not in sekarang:
                            sekarang.append(item)
                    baru[kunci] = sekarang
                elif isinstance(nilai, ArrayRemove):
                    sekarang = list(baru.get(kunci) or [])
                    baru[kunci] = [x for x in sekarang if x not in nilai.values]
                elif nilai is SERVER_TIMESTAMP:
                    baru[kunci] = _dt.datetime.now()
                else:
                    baru[kunci] = nilai
            self._tulis(baru, conn)
            if luar:
                conn.commit()
        except Exception:
            if luar:
                conn.rollback()
            raise
        return self

    def delete(self):
        with _koneksi().cursor() as cur:
            cur.execute(f"DELETE FROM `{self._tabel}` WHERE `id`=%s", (self._id,))
        return None

    def __repr__(self):
        return f"<DocumentReference {self._tabel}/{self._id}>"


class Query:
    def __init__(self, tabel, filter_list=None, batas=None):
        self._tabel = tabel
        self._filter = list(filter_list or [])
        self._batas = batas

    def where(self, field, op, value):
        return Query(self._tabel, self._filter + [(field, op, value)], self._batas)

    def limit(self, jumlah):
        return Query(self._tabel, self._filter, int(jumlah))

    def _sql(self):
        where, params = _susun_where(self._tabel, self._filter)
        sql = f"SELECT `id`, `data` FROM `{self._tabel}`{where} ORDER BY `id` ASC"
        if self._batas is not None:
            sql += " LIMIT %s"
            params.append(self._batas)
        return sql, params

    def stream(self):
        sql, params = self._sql()
        with _koneksi().cursor() as cur:
            cur.execute(sql, params)
            for baris in cur.fetchall():
                data = _load_json(baris["data"])
                ref = DocumentReference(self._tabel, baris["id"])
                yield DocumentSnapshot(ref, data)

    def get(self):
        return list(self.stream())

    def __iter__(self):
        return self.stream()


class CollectionReference:
    def __init__(self, tabel):
        self._tabel = tabel

    @property
    def id(self):
        return self._tabel

    def document(self, doc_id=None):
        if doc_id is None:
            doc_id = "".join(
                _secrets.choice(_string.ascii_letters + _string.digits)
                for _ in range(20)
            )
        return DocumentReference(self._tabel, str(doc_id))

    def where(self, field, op, value):
        return Query(self._tabel).where(field, op, value)

    def limit(self, jumlah):
        return Query(self._tabel).limit(jumlah)

    def stream(self):
        return Query(self._tabel).stream()

    def get(self):
        return Query(self._tabel).get()

    def __repr__(self):
        return f"<CollectionReference {self._tabel}>"


class Transaction:
    """Transaksi MySQL (pengganti firestore.transaction)."""

    def __init__(self):
        self._conn = _koneksi_baru()
        self._conn.autocommit(False)

    def set(self, reference, data):
        reference._tulis(dict(data or {}), self._conn)
        return self

    def commit(self):
        self._conn.commit()

    def rollback(self):
        self._conn.rollback()

    def close(self):
        try:
            self._conn.close()
        except Exception:
            pass


def transactional(fungsi):
    """Dekorator pengganti @firestore.transactional."""

    @_functools.wraps(fungsi)
    def pembungkus(transaction=None):
        tx = transaction if transaction is not None else Transaction()
        try:
            hasil = fungsi(tx)
            tx.commit()
            return hasil
        except Exception:
            tx.rollback()
            raise
        finally:
            tx.close()

    return pembungkus


class Client:
    def collection(self, nama):
        if nama not in TABEL:
            raise ValueError(
                f"Koleksi '{nama}' tidak dikenal. "
                f"Tambahkan dulu di TABEL (db_mysql.py)."
            )
        return CollectionReference(nama)

    def transaction(self):
        return Transaction()

    def siapkan_skema(self):
        siapkan_skema()


_klien = None
_kunci_klien = _threading.Lock()


def get_db():
    """Ambil objek database (pengganti `firestore.client()`)."""
    global _klien
    if _klien is None:
        with _kunci_klien:
            if _klien is None:
                try:
                    siapkan_skema()
                except Exception as exc:  # jangan gagalkan start aplikasi
                    print(f"[PERINGATAN] Gagal menyiapkan MySQL: {exc}")
                    print(
                        "[PERINGATAN] Pastikan MySQL/XAMPP berjalan dan "
                        "konfigurasi MYSQL_* benar."
                    )
                _klien = Client()
    return _klien


def info_koneksi():
    """Ringkasan konfigurasi tanpa menampilkan password (untuk diagnostik)."""
    return {
        "host": KONFIG["host"],
        "port": KONFIG["port"],
        "user": KONFIG["user"],
        "database": KONFIG["database"],
    }

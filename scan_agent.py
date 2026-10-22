"""
Scan Agent lokal untuk Portal Karyawan.

Dijalankan di komputer pengguna yang memiliki scanner USB (mis. Epson L3210
atau DS-410). Browser akan memanggil http://127.0.0.1:3000, bukan server portal.

Jalankan:
    python scan_agent.py

Dependensi:
    pip install -r requirements-scan-agent.txt
"""

import os
import tempfile
import threading
import time
import uuid

from flask import Flask, jsonify, request, send_from_directory
from flask_cors import CORS

try:
    from waitress import serve
except ImportError:  # fallback untuk development
    serve = None

app = Flask(__name__)
# Agent hanya mendengarkan di localhost. CORS longgar agar portal LAN bisa
# memanggilnya; jangan pernah membuka port ini ke jaringan.
CORS(app, resources={r"/api/*": {"origins": "*"}})

SCAN_LOCK = threading.Lock()
SCAN_DIR = os.path.join(tempfile.gettempdir(), "portal_scan_agent")
os.makedirs(SCAN_DIR, exist_ok=True)


@app.after_request
def _header_private_network(resp):
    # Diperlukan Chrome saat portal HTTPS memanggil agent localhost.
    resp.headers["Access-Control-Allow-Private-Network"] = "true"
    return resp

FORMAT_PNG = "{B96B3CAF-0728-11D3-9D7B-0000F81EF32E}"


def _com_init():
    try:
        import pythoncom
        pythoncom.CoInitialize()
    except Exception:
        pass


def _com_uninit():
    try:
        import pythoncom
        pythoncom.CoUninitialize()
    except Exception:
        pass


def _wia_devices():
    import win32com.client
    dm = win32com.client.Dispatch("WIA.DeviceManager")
    hasil = []
    for i in range(1, dm.DeviceInfos.Count + 1):
        try:
            info = dm.DeviceInfos(i)
            props = {}
            for nama in ("Name", "Description", "Port"):
                try:
                    props[nama] = info.Properties(nama).Value
                except Exception:
                    props[nama] = ""
            hasil.append({
                "index": i,
                "id": info.DeviceID,
                "nama": props.get("Name") or "",
                "deskripsi": props.get("Description") or "",
            })
        except Exception:
            continue
    return hasil


def _bersihkan_lama():
    batas = time.time() - 3600
    for nama in os.listdir(SCAN_DIR):
        path = os.path.join(SCAN_DIR, nama)
        try:
            if os.path.isfile(path) and os.path.getmtime(path) < batas:
                os.remove(path)
        except OSError:
            pass


@app.get("/api/status")
def status():
    _com_init()
    try:
        devs = _wia_devices()
        return jsonify({"online": bool(devs), "jumlah": len(devs), "perangkat": devs})
    except Exception as e:
        return jsonify({"online": False, "jumlah": 0, "perangkat": [], "error": str(e)})
    finally:
        _com_uninit()


@app.post("/api/scan")
def scan():
    data = request.get_json(silent=True) or {}
    try:
        dpi = int(data.get("dpi") or 200)
    except Exception:
        dpi = 200
    dpi = min(max(dpi, 75), 600)
    mode_warna = (data.get("colorMode") or "color").strip().lower()
    duplex = bool(data.get("duplex"))
    try:
        maks = int(data.get("maxPages") or 0)
    except Exception:
        maks = 0
    maks = maks if 0 < maks <= 50 else 50
    prefix = "".join(c for c in str(data.get("prefix") or "scan") if c.isalnum() or c in ("-", "_"))[:40] or "scan"

    try:
        index = int(data.get("deviceIndex") or 1)
    except Exception:
        index = 1

    _com_init()
    tmpdir = tempfile.mkdtemp(prefix="scan_agent_")
    pngs = []
    try:
        with SCAN_LOCK:
            import fitz
            import win32com.client

            _bersihkan_lama()
            dm = win32com.client.Dispatch("WIA.DeviceManager")
            if dm.DeviceInfos.Count < 1:
                return jsonify({
                    "success": False,
                    "error": "Tidak ada scanner terdeteksi di komputer ini. Pastikan scanner menyala, USB terpasang, dan driver terinstal."
                }), 404

            idx = index if 1 <= index <= dm.DeviceInfos.Count else 1
            dev = dm.DeviceInfos(idx).Connect()

            ada_feeder = False
            try:
                dev.Properties("Document Handling Select").Value = 1 | (4 if duplex else 0)
                ada_feeder = True
            except Exception:
                pass

            item = dev.Items(1)
            try:
                item.Properties("Horizontal Resolution").Value = dpi
                item.Properties("Vertical Resolution").Value = dpi
            except Exception:
                pass
            try:
                item.Properties("Pages").Value = 1
            except Exception:
                pass
            try:
                item.Properties("Current Intent").Value = {"color": 1, "grayscale": 2, "bw": 4}.get(mode_warna, 1)
            except Exception:
                pass

            def feeder_siap():
                try:
                    return bool(dev.Properties("Document Handling Status").Value & 1)
                except Exception:
                    return True

            def transfer_png():
                try:
                    return item.Transfer()
                except Exception:
                    dlg = win32com.client.Dispatch("WIA.CommonDialog")
                    return dlg.ShowTransfer(item, FORMAT_PNG, False)

            n = 0
            while n < maks:
                if n > 0 and not feeder_siap():
                    break
                if n >= 1 and not ada_feeder:
                    break
                try:
                    img = transfer_png()
                except Exception as e:
                    if n == 0:
                        return jsonify({
                            "success": False,
                            "error": f"Gagal memindai. Pastikan dokumen sudah diletakkan di scanner: {e}"
                        }), 500
                    break
                fp = os.path.join(tmpdir, f"p{n + 1}.png")
                try:
                    img.SaveFile(fp)
                except Exception:
                    break
                pngs.append(fp)
                n += 1

            if not pngs:
                return jsonify({
                    "success": False,
                    "error": "Tidak ada halaman yang berhasil dipindai. Letakkan dokumen di scanner lalu coba lagi."
                }), 400

            pdf = fitz.open()
            for fp in pngs:
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
                    page.insert_image(page.rect, stream=pix.tobytes("jpg", jpg_quality=80))
                finally:
                    pix = None

            nama = f"{prefix}_{uuid.uuid4().hex[:8]}.pdf"
            out = os.path.join(SCAN_DIR, nama)
            pdf.save(out)
            pdf.close()

            return jsonify({
                "success": True,
                "document": {
                    "downloadUrl": f"/api/scan/download/{nama}",
                    "pageCount": len(pngs),
                    "filename": nama,
                }
            })
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
            import shutil
            shutil.rmtree(tmpdir, ignore_errors=True)
        except Exception:
            pass
        _com_uninit()


@app.get("/api/scan/download/<path:nama>")
def download(nama):
    safe = os.path.basename(nama)
    return send_from_directory(SCAN_DIR, safe, as_attachment=False, mimetype="application/pdf")


if __name__ == "__main__":
    print("Scan Agent lokal berjalan di http://127.0.0.1:3000")
    print("Jangan tutup jendela ini saat memakai fitur scan.")
    if serve:
        serve(app, host="127.0.0.1", port=3000, threads=4)
    else:
        app.run(host="127.0.0.1", port=3000, threaded=True)

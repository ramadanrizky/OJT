# Scan Agent Lokal

Dipakai jika scanner USB terpasang di **komputer pengguna**, bukan di server.

## Cara cepat (tanpa Python)

1. Salin folder `scan-agent-dist` (berisi `ScanAgent.exe` dan `JALANKAN-SCAN-AGENT.bat`) ke komputer yang memiliki scanner.
2. Buka folder tersebut, lalu klik dua kali `JALANKAN-SCAN-AGENT.bat`.
3. Biarkan jendela agent terbuka, lalu buka portal dan lakukan scan.

Agent hanya mendengarkan di komputer itu:

```text
http://127.0.0.1:3000
```

## Syarat

- Windows
- Scanner + driver WIA/TWAIN sudah terpasang dan terdeteksi Windows
- Komputer dan portal berada di jaringan yang sama

## Menjalankan dari source Python

Jika memakai `scan_agent.py`, instal dulu dependensinya:

```powershell
pip install -r requirements-scan-agent.txt
```

Lalu klik dua kali `JALANKAN-SCAN-AGENT.bat`, atau:

```powershell
python scan_agent.py
```

## Agar otomatis jalan saat login

1. Tekan `Win + R`, ketik `shell:startup`, Enter.
2. Buat shortcut `JALANKAN-SCAN-AGENT.bat` di folder Startup tersebut.

## Cek scanner terdeteksi

Buka browser di komputer yang sama:

```text
http://127.0.0.1:3000/api/status
```

Contoh hasil normal:

```json
{"online": true, "jumlah": 1, "perangkat": [{"nama": "Epson L3210", "index": 1}]}
```

## Catatan

- Agent ini hanya untuk Windows. Untuk HP/tablet, gunakan pemindai kamera pada halaman Scan Dokumen.
- Browser tidak bisa mengakses scanner USB secara langsung. Karena itu agent ini wajib berjalan di komputer yang memiliki scanner.
- Jika portal dibuka lewat HTTPS, browser tertentu dapat memblokir panggilan ke `http://127.0.0.1`. Untuk jaringan internal, gunakan HTTP portal atau akses portal lewat alamat LAN.
- Halaman scan tetap mencoba scanner server sebagai cadangan jika agent lokal tidak aktif.
- Scanner yang terpasang di satu komputer tidak otomatis bisa dipakai komputer lain. Jika semua orang harus memakai scanner itu, pindahkan scanner ke server atau jadikan komputer tersebut sebagai scan station.
- `ScanAgent.exe` belum ditandatangani digital. Windows SmartScreen dapat menampilkan peringatan saat pertama dijalankan; untuk distribusi resmi sebaiknya exe ditandatangani.

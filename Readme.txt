set FLASK_APP=app.py
flask run


 cloudflared.exe tunnel --url http://localhost:5000

apakah scan file 
 user id pada profile caranya : @userinfobot


restore database ke server baru

Pastikan MySQL/XAMPP sudah jalan di server baru.
Hentikan portal.
Jalankan:
"C:\xampp\mysql\bin\mysql.exe" -u root -p < "D:\Backup-Portal\portal_karyawan_2026-10-21.sql"
Kalau password root kosong, tekan Enter.
Copy folder uploads dan outputs ke server baru.
Sesuaikan konfigurasi database di JALANKAN-APP.bat:
set MYSQL_USER=satsibertni
set MYSQL_PASSWORD=satsiber2026
set MYSQL_DB=portal_karyawan
Kalau database ada di server lain, tambahkan:
set MYSQL_HOST=IP-SERVER-DATABASE
Jalankan portal dan tes login serta buka data lama.



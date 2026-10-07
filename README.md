# Loading Control Tower

Aplikasi untuk mencatat urutan datang dan urutan muat mobil di area loading. Tampilannya (halaman web) ada di GitHub Pages, datanya di Google Sheets, dan logikanya di Google Apps Script.

## Isi repositori

```
loading-control-tower/
├── README.md                     panduan ini
├── docs/                         halaman web (yang dibuka petugas)
│   ├── index.html
│   ├── security.html
│   ├── juru.html
│   ├── checker.html
│   ├── dashboard.html
│   ├── config.js                 satu-satunya file yang perlu diedit
│   ├── app.js
│   ├── style.css
│   ├── sw.js
│   ├── manifest.webmanifest
│   └── icon-192.png, icon-512.png, apple-touch-icon.png
├── apps-script/                  kode server (ditempel ke Apps Script)
│   ├── Code.gs
│   └── appsscript.json
├── data/
│   └── Database_Loading_Control_Tower.xlsx
└── panduan/
    └── Panduan_Penggunaan_Loading_Control_Tower.pdf
```

## Cara pasang

### 1. Google Sheets dan Apps Script
1. Unggah `data/Database_Loading_Control_Tower.xlsx` ke Google Drive, buka dengan Google Sheets, lalu File > Save as Google Sheets.
2. File > Settings, atur zona waktu ke Jakarta (GMT+07:00).
3. Extensions > Apps Script. Hapus isi `Code.gs` bawaan, lalu tempel isi `apps-script/Code.gs`.
4. Project Settings, centang Show "appsscript.json" manifest file. Buka file itu di editor dan ganti isinya dengan `apps-script/appsscript.json`.
5. Muat ulang Google Sheets, lalu pilih menu Control Tower > Setup tab otomatis.
6. Di tab USER, isi nama petugas dengan peran SECURITY, JURU_MUAT, atau CHECKER, dan Aktif = Ya.
7. Deploy > New deployment > Web app. Execute as: Me. Who has access: Anyone. Tekan Deploy, izinkan akses, lalu salin alamat Web app (berakhiran `/exec`).

### 2. Isi alamat di config.js
Buka `docs/config.js` dan ganti `PASTE_ALAMAT_WEB_APP_DI_SINI` dengan alamat tadi. Hasil akhirnya kira-kira begini:

```
window.API_URL='https://script.google.com/macros/s/xxxx/exec';
```

### 3. Unggah ke GitHub
1. Buat repositori baru di github.com (Public).
2. Add file > Upload files. Seret seluruh isi folder `loading-control-tower` (termasuk folder `docs`, `apps-script`, `data`, `panduan`), lalu Commit changes.
3. Settings > Pages > Source: Deploy from a branch. Branch: main, folder: /docs. Save.
4. Tunggu 1 sampai 2 menit.

### 4. Alamat yang dibagikan
Ganti NAMA dan REPO dengan milik Anda.

```
https://NAMA.github.io/REPO/security.html
https://NAMA.github.io/REPO/juru.html
https://NAMA.github.io/REPO/checker.html
https://NAMA.github.io/REPO/dashboard.html
```

Menu semua halaman: `https://NAMA.github.io/REPO/?menu`

## Cek koneksi
Buka `ALAMAT_WEB_APP?fn=apiView&arg=%22dashboard%22` di browser. Yang muncul harus teks JSON berawalan `{"ok":true`. Kalau tidak, deploy Apps Script belum benar (ulangi langkah 7).

## Memperbarui
- Mengubah `Code.gs`: tempel ulang di Apps Script, lalu Deploy > Manage deployments > pensil > Version: New version > Deploy.
- Mengubah file di `docs/`: unggah ulang file itu ke GitHub. Kalau tampilan lama masih muncul di HP, ubah `ct-v1` di `docs/sw.js` menjadi `ct-v2`.

## Catatan
Isi repositori publik bisa dilihat siapa saja, termasuk alamat Web app. Jangan simpan data rahasia di sini. Data mobil tetap aman di Google Sheets Anda.

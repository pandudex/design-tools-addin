# Design Tools (PowerPoint Add-in, prototype)

Panel di PowerPoint berisi 25 fitur dari daftar fitur:
Align (6), Dock (4), Stretch (4), Fill (4), Distribute (2), Slice/Multiply, Align in Matrix,
Split/Merge Text Box, Rounded Rectangle Radius.

## Cara pakai
- **Master** = objek yang dipilih terakhir. Pilih mode lain di dropdown Master jika urutan seleksi tidak sesuai.
- **Pilih 1 objek, atau tahan Ctrl saat klik** = terhadap slide, bukan Master.
- Distribute: min. 3 objek (2 objek jika Ctrl, disebar di seluruh slide).

## Menjalankan
Butuh Node.js 18+ dan PowerPoint (Windows/Mac/web).

```bash
npm install
npm run certs      # sekali saja
npm run serve      # biarkan terbuka (https://localhost:3000)
npm run sideload   # terminal lain; PowerPoint web: upload manifest.xml
npm test           # uji logika geometri
```
Lalu di PowerPoint: tab **Home** > **Design Tools**.

## Struktur (mudah ditambah fitur)
- `src/geometry.js`: logika murni (rumus posisi/ukuran), diuji di `tests/`
- `src/taskpane.js`: penghubung ke Office.js. Fitur baru = 1 fungsi di objek `ACTIONS` + tombol `data-act` di HTML
- `src/icons/`: 25 ikon SVG fitur (nama file = nama fitur, mis. `dock-up.svg`). Ganti file dengan nama sama untuk memperbarui ikon
- `assets/icon-16|32|64|80.png`: ikon tombol di pita PowerPoint (masih sementara)
- `manifest.xml`, `src/taskpane.html|css`

## Batasan API PowerPoint (perlu kamu ketahui)
1. **Urutan seleksi** tidak dijamin dokumentasi Office.js. Master "dipilih terakhir" dan urutan matriks/merge mengandalkan urutan yang dikembalikan PowerPoint; belum diuji di PowerPoint asli.
2. **Slice/Multiply**: Office.js tidak punya "duplicate shape". Salinan dibuat ulang sebagai persegi (atau rounded rectangle jika namanya mengandung "Round") dengan warna isi dan garis yang sama. Teks, gradasi, dan efek tidak ikut tersalin.
3. **Rounded Rectangle Radius** dan ukuran slide butuh PowerPointApi 1.10 (Microsoft 365 versi baru). Nilai adjustment diasumsikan pecahan 0-0.5 dari sisi terpendek. Shape dikenali dari namanya ("Round").
4. **Split Text Box**: kotak kedua ditaruh tepat di bawah kotak asli; font nama/ukuran/warna/tebal/miring disalin, format per-kata tidak.
5. Ctrl+klik di panel dipakai sebagai pengganti "tekan Ctrl" pada spesifikasi.

## Sebelum dibagikan ke klien
Ganti `https://localhost:3000` di manifest dengan URL HTTPS asli lalu distribusikan lewat AppSource atau admin center Microsoft 365.

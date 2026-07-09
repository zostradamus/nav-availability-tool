# NAV 2G Availability Analysis Tool

Pipeline otomatis untuk menganalisis **network availability (NAV) 2G** per site di region Central Java (IOH) dari dataset mingguan, dengan pengayaan lokasi sampai level **kelurahan** dan atribut infrastruktur dari **CMDB**.

Dibuat saat program magang di Indosat Ooredoo Hutchison.

## Fitur

- **Deteksi otomatis dataset baru** — scan folder NAV mingguan (`Week_<N>_<tahun>/*.xlsb`) dan file CMDB (`*.xlsx`, dipilih nomor W terbesar); hanya file baru/berubah yang diproses (dilacak via signature nama+ukuran+mtime).
- **Availability rata-rata 7 hari** per site dari kolom NAV harian, dikategorikan: `100%`, `98-<100%`, `<98%`.
- **Reverse-geocode kelurahan** — koordinat site dicocokkan point-in-polygon dengan batas 8.728 kelurahan Jateng+DIY (hasil di-cache per site).
- **Join CMDB** via Site ID: site type, transport, priority, VIP, genset, battery backup, owner, alamat. Sel kosong diberi kode (`NOT_IN_CMDB` / `BLANK_IN_CMDB`).
- **Output Excel** multi-sheet: Summary (+legenda warna & kode), Data mingguan (formula hidup + pewarnaan harian), 3 sheet kategori, Top 10 worst, No_NAV_Data (site tanpa data monitoring), Weekly_Trend.
- **Portabel & config-driven** — semua path diatur di `config.json`; `copy_to` untuk menyalin hasil otomatis ke folder lain (mis. folder OneDrive tersinkron).
- **Tahan banting** — file output terkunci? tersimpan sebagai `*_new.xlsx`. Config pakai backslash Windows? dikoreksi otomatis. Dependensi Python ter-install sendiri di run pertama.

## Cara Pakai

```bash
# 1. salin config
cp config.example.json config.json   # lalu edit path-nya

# 2. jalankan
python run_nav_analysis.py               # proses file baru + generate Excel
python run_nav_analysis.py --no-excel    # proses file baru saja
python run_nav_analysis.py --excel-only  # generate Excel dari cache
python run_nav_analysis.py --force       # proses ulang semuanya
```

Di Windows cukup double-click `run.bat` (output tercatat ke `nav_log.txt`).

## Konfigurasi (`config.json`)

| Kunci | Isi |
|---|---|
| `nav_dir` | Folder berisi subfolder `Week_<N>_<tahun>/` dengan file `.xlsb` NAV |
| `cmdb_dir` | Folder berisi file CMDB `.xlsx` (otomatis dipakai nomor W terbesar) |
| `output_xlsx` | Nama/path file Excel hasil |
| `region` | Nilai kolom `New_Region` yang difilter, mis. `CENTRAL JAVA` |
| `copy_to` | Daftar folder tujuan salinan hasil (mis. folder OneDrive sync) |

Path relatif dihitung dari folder tool; gunakan `/` atau `\\` (backslash tunggal juga ditoleransi).

## Penjadwalan Otomatis (Windows)

Buat task harian di **Task Scheduler**: Create Basic Task -> Daily -> Start a program -> arahkan ke `run.bat` (kolom *Start in* diisi folder tool **tanpa tanda kutip**).

```bat
:: jalankan sekarang
schtasks /run /tn "NAV 2G Central Java Update"
:: cek status (Last Result 0 = sukses)
schtasks /query /tn "NAV 2G Central Java Update" /v /fo list | findstr /i "Result Status"
:: hentikan paksa
schtasks /end /tn "NAV 2G Central Java Update"
```

## Struktur

```
run_nav_analysis.py   # pipeline utama
run.bat               # wrapper Windows (log ke nav_log.txt)
config.example.json   # template konfigurasi
ref/                  # batas kelurahan Jateng+DIY (pickle WKB)
cache/                # dibuat otomatis: state, parquet mingguan, cache kelurahan
```

`cache/` aman dihapus - dibangun ulang otomatis (run berikutnya lebih lama sekali saja).

## Kredit & Lisensi Data

Batas kelurahan diolah dari [cahyadsn/wilayah_boundaries](https://github.com/cahyadsn/wilayah_boundaries) (MIT License).

> **Catatan:** repo ini hanya berisi kode dan data referensi publik. Dataset NAV/CMDB serta hasil analisis adalah data internal perusahaan dan tidak disertakan (lihat `.gitignore`).


DASHBOARD INTERAKTIF
  Setiap run juga menghasilkan dashboard.html - buka di browser.
  Fitur: ringkasan + trend, rekap per kecamatan (klik untuk detail site),
  site explorer (filter kategori/kabupaten/pencarian, export CSV),
  top 10 worst, analisis bulanan (site konsisten <98%), perbandingan
  antar minggu, mode 2G / 4G / traffic, dan detail CMDB per site.
  Kunci config: output_dashboard (default dashboard.html), ikut tersalin
  ke folder copy_to.

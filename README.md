# `indo-terbilang` (Excel Macro / VBA)

Modul **Excel Macro / VBA (`.bas`)** tanpa dependensi eksternal (*zero dependencies*) untuk mengonversi angka menjadi kalimat **terbilang Bahasa Indonesia baku (EYD V / KBBI)** serta memformat dan mem-parsing mata uang **Rupiah (IDR)** langsung di rumus sel Excel maupun Macro VBA.

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](./LICENSE)
[![Zero Dependencies](https://img.shields.io/badge/dependencies-0-brightgreen.svg)](./src/IndoTerbilang.bas)
[![Platform: Excel Windows & Mac](https://img.shields.io/badge/Excel-Windows%20%26%20Mac-217346.svg)](#cara-memasang-di-excel)

---

## Fitur Utama

- **100% Pure VBA (Zero Dependencies):** Tidak menggunakan `CreateObject`, ActiveX, atau `VBScript.RegExp` — berjalan mulus di **Excel Windows (32-bit & 64-bit)** maupun **Excel for Mac** (Excel 2007 s.d. Microsoft 365).
- **Bebas Masalah Regional Setting Windows:** Menggunakan konversi angka invarian (`Str$`), sehingga rumus tetap konsisten baik di Windows setting **Indonesia** (`,` desimal) maupun **English (US)** (`.` desimal).
- **Langsung Jadi Rumus Excel (UDF) & Fungsi Macro:**
  - `=Terbilang(...)`
  - `=TerbilangRupiah(...)`
  - `=FormatRupiah(...)`
  - `=ParseRupiah(...)`
- **Presisi Hingga 36 Digit (`desiliun`):** Mendukung input angka biasa, tipe `Currency` / `Decimal`, maupun teks angka super panjang hingga $10^{36} - 1$.
- **Akurat EYD V / KBBI:** `seribu`, `sebelas`, `seratus`, `satu juta`, ejaan baku `miliar`, pembacaan desimal digit-per-digit (`3.14` -> `"tiga koma satu empat"`), serta penanganan `sen` dengan pembulatan berbasis string (`half-up` / `truncate`).
- **Fail-Loud yang Aman di Excel & VBA:**
  - Jika dipanggil dari **sel Excel (UDF)**: input tidak valid mengembalikan `#VALUE!` (`INVALID_INPUT`), sedangkan angka non-finite / melebihi 36 digit mengembalikan `#NUM!` (`OUT_OF_RANGE`), sehingga aman dibungkus `=IFERROR(...)`.
  - Jika dipanggil dari **kode Macro VBA (`Sub`)**: melempar error terstruktur (`TERBILANG_ERR_INVALID_INPUT`, `TERBILANG_ERR_NON_FINITE`, `TERBILANG_ERR_OUT_OF_RANGE`) yang dapat ditangkap via `On Error GoTo`.

> **Varian Bahasa Lain:**
> - **TypeScript / JavaScript (npm):** https://github.com/salmankalista/indo-terbilang
> - **Python:** https://github.com/salmankalista/indo-terbilang-python
> - **PHP:** https://github.com/salmankalista/indo-terbilang-php

---

## Unduh Repository

- **Download ZIP (Paling Mudah):** Unduh arsip terbaru dari https://github.com/salmankalista/indo-terbilang-excel-macro/archive/refs/heads/main.zip (atau klik **`Code` -> `Download ZIP`** di GitHub), lalu ekstrak.
- **Atau Clone via Git:**
  ```bash
  git clone https://github.com/salmankalista/indo-terbilang-excel-macro.git
  ```
- **Butuh File Modulnya Saja?** File utama yang perlu diimpor ke Excel hanya satu file: [`src/IndoTerbilang.bas`](./src/IndoTerbilang.bas).

---

## Cara Memasang di Excel

### Cara 1: Import File `.bas` ke File Excel (Paling Praktis)

1. Buka file Excel Anda, lalu tekan **`Alt + F11`** (Mac: **`Fn + Option + F11`**) untuk membuka **VBA Editor**.
2. Pada menu di atas, klik **`File` -> `Import File...`** (atau tekan **`Ctrl + M`**).
3. Pilih file [`src/IndoTerbilang.bas`](./src/IndoTerbilang.bas), lalu klik **Open**.
4. Tutup VBA Editor (kembali ke jendela Excel).
5. Simpan file Anda dengan format **Excel Macro-Enabled Workbook (`.xlsm`)** atau **Excel Binary Workbook (`.xlsb`)** agar macro tidak hilang (jangan disimpan sebagai `.xlsx`).
6. Sekarang Anda sudah bisa memakai rumus `=Terbilang(...)` dan `=TerbilangRupiah(...)` langsung di sel Excel!

> **Tips:** Jika ingin deskripsi fungsi & keterangan parameter muncul rapi di kotak dialog **Insert Function (`fx`)**, tekan **`Alt + F8`**, ketik/pilih **`RegisterIndoTerbilangFunctions`**, lalu klik **Run** (cukup 1 kali per file/Add-In).

### Cara 2: Pasang Permanen sebagai Excel Add-In (`.xlam`)

Gunakan cara ini jika Anda ingin rumus `=Terbilang(...)` otomatis tersedia di **semua file Excel** (termasuk file `.xlsx` biasa tanpa perlu import berulang):

1. Buka workbook baru kosong di Excel.
2. Tekan **`Alt + F11`** -> **`File` -> `Import File...`** -> pilih [`src/IndoTerbilang.bas`](./src/IndoTerbilang.bas).
3. *(Opsional)* Jalankan `RegisterIndoTerbilangFunctions` lewat **`Alt + F8`** atau Immediate Window (`Ctrl + G`).
4. Simpan file melalui **`File` -> `Save As`**, ubah *Save as type* menjadi **`Excel Add-In (*.xlam)`** (Excel akan otomatis mengarahkan ke folder `AddIns`), lalu simpan dengan nama misalnya `IndoTerbilang.xlam`.
5. Aktifkan Add-In melalui **`File` -> `Options` -> `Add-ins` -> `Manage: Excel Add-ins` -> `Go...`**, centang **IndoTerbilang**, lalu klik **OK**.

---

## Quick Start (Rumus di Sel Excel)

> **Penting — Pemisah Argumen Rumus Excel (`;` vs `,`):**
> - Pada Windows dengan Regional Setting **Indonesia**, Excel memakai titik koma **`;`** (contoh: `=Terbilang(A1; "title")`).
> - Pada Windows/Mac dengan Regional Setting **English (US)**, Excel memakai koma **`,`** (contoh: `=Terbilang(A1, "title")`).

| Rumus (Setting Indonesia `;`) | Contoh Nilai `A1` | Hasil |
| :--- | :--- | :--- |
| `=Terbilang(A1)` | `1500000` | `satu juta lima ratus ribu` |
| `=Terbilang(A1; "title"; "Terbilang:"; "rupiah")` | `1500000` | `Terbilang: Satu Juta Lima Ratus Ribu Rupiah` |
| `=Terbilang(A1)` | `3.14` | `tiga koma satu empat` |
| `=TerbilangRupiah(A1)` | `1250.5` | `seribu dua ratus lima puluh rupiah lima puluh sen` |
| `=TerbilangRupiah(A1; "title")` | `1250750.5` | `Satu Juta Dua Ratus Lima Puluh Ribu Tujuh Ratus Lima Puluh Rupiah Lima Puluh Sen` |
| `=TerbilangRupiah(A1; "sentence"; ; ; ; ; FALSE)` | `1250.6` | `Seribu dua ratus lima puluh satu rupiah` |
| `=FormatRupiah(A1)` | `1500000` | `Rp 1.500.000,00` |
| `=FormatRupiah(A1; ; ; ; ; ; ; ; TRUE)` | `1500000` | `Rp 1,5 jt` |
| `=FormatRupiah(A1; ; ; ; ; 0; ; "parentheses")` | `-750000` | `(Rp 750.000)` |
| `=ParseRupiah(A1)` | `Rp 1.500.000,50` | `1500000.50` |

---

## Tutorial Pemakaian Praktis

### Tutorial 1: Rumus Kwitansi / Invoice Otomatis (Aman untuk Sel Kosong)

Jika kolom total tagihan berada di sel `E12` dan ingin tetap kosong saat `E12` belum diisi:

```excel
=IF(E12=""; ""; TerbilangRupiah(E12; "title"))
```

Atau gunakan `IFERROR` agar sel tidak menampilkan `#VALUE!` saat data belum valid:

```excel
=IFERROR(TerbilangRupiah(E12; "sentence"; "Terbilang:"); "")
```

### Tutorial 2: Menangani Nominal Raksasa (> 15 Digit) di Excel

Excel secara bawaan membulatkan angka numerik setelah **15 digit signifikan**. Untuk nominal sangat besar hingga 36 digit (`desiliun`), simpan sel sebagai **Teks** (awali tanda petik satu `'` saat mengetik, atau ubah Number Format sel menjadi **Text (`@`)**):

- Sel `A1` (format Text): `12345678901234567890.75`
- Rumus `=TerbilangRupiah(A1; "title")` ->
  `"Dua Belas Kuintiliun Tiga Ratus Empat Puluh Lima Kuadriliun Enam Ratus Tujuh Puluh Delapan Triliun Sembilan Ratus Satu Miliar Dua Ratus Tiga Puluh Empat Juta Lima Ratus Enam Puluh Tujuh Ribu Delapan Ratus Sembilan Puluh Rupiah Tujuh Puluh Lima Sen"`

### Tutorial 3: Mengurai Teks `"Rp 1.500.000,50"` dari Hasil Copy-Paste PDF/Web

Jika sel `A1` berisi teks mentah `"Rp 1.500.000,50"` (bukan angka Excel):

```excel
=TerbilangRupiah(ParseRupiah(A1); "title")
```

Dan jika ingin mengubahnya menjadi angka Excel agar bisa dijumlahkan (`SUM`) secara aman di semua pengaturan bahasa Windows:

```excel
=NUMBERVALUE(ParseRupiah(A1); ".")
```

### Tutorial 4: Menggunakan Macro Contoh & Demo Playground Otomatis

Impor juga [`examples/IndoTerbilangExamples.bas`](./examples/IndoTerbilangExamples.bas), lalu tekan **`Alt + F8`** dan jalankan:

- **`BuatSheetDemoIndoTerbilang`**: Membuat worksheet `Demo_IndoTerbilang` lengkap berisi contoh rumus `Terbilang`, `TerbilangRupiah`, `FormatRupiah`, dan `ParseRupiah`.
- **`IsiKolomTerbilangDariSeleksi`**: Sorot sel/kolom berisi angka (atau teks `"Rp ..."`), jalankan macro ini, dan kolom di sebelah kanannya otomatis terisi kalimat `TerbilangRupiah` (Title Case).
- **`DemoTerbilangQuickStart`**: Mencetak contoh pemanggilan VBA ke Immediate Window (`Ctrl + G`).

### Tutorial 5: Menangkap Error di Macro VBA (`Sub`)

```vba
Sub ProsesNominal()
    On Error GoTo ErrHandler
    Dim hasil As String
    hasil = TerbilangRupiah(ParseRupiah("Rp 2.500.000,75"), "sentence")
    MsgBox hasil, vbInformation
    Exit Sub

ErrHandler:
    Select Case Err.Number
        Case TERBILANG_ERR_INVALID_INPUT
            MsgBox "Input salah: " & Err.Description, vbExclamation
        Case TERBILANG_ERR_OUT_OF_RANGE
            MsgBox "Angka melebihi batas 36 digit: " & Err.Description, vbExclamation
        Case Else
            MsgBox "Error: " & Err.Description, vbCritical
    End Select
End Sub
```

---

## Referensi API

### 1. `Terbilang(Value, [Style], [Prefix], [Suffix], [NegativePrefix], [DecimalWord], [TrimTrailingDecimalZeros])`

| Parameter | Tipe | Default | Keterangan |
| :--- | :--- | :--- | :--- |
| `Value` | `Variant` | *(wajib)* | Angka, `Range` sel tunggal, atau teks numerik (maks. 36 digit bulat). |
| `Style` | `String` | `"lower"` | `"lower"`, `"upper"`, `"title"`, atau `"sentence"`. |
| `Prefix` | `String` | `""` | Teks awalan sebelum hasil. |
| `Suffix` | `String` | `""` | Teks akhiran setelah hasil. |
| `NegativePrefix` | `String` | `"minus"` | Kata pengganti tanda `-`. |
| `DecimalWord` | `String` | `"koma"` | Kata pemisah angka desimal. |
| `TrimTrailingDecimalZeros` | `Boolean` | `False` | `True` untuk memangkas nol di ujung kanan desimal teks (mis. `"12.50"` -> `"dua belas koma lima"`). |

### 2. `TerbilangRupiah(Value, [Style], [Prefix], [Suffix], [NegativePrefix], [CurrencyUnit], [IncludeSen], [SenUnit], [Rounding])`

| Parameter | Tipe | Default | Keterangan |
| :--- | :--- | :--- | :--- |
| `Value` | `Variant` | *(wajib)* | Angka, `Range` sel tunggal, atau teks numerik. |
| `Style` | `String` | `"lower"` | `"lower"`, `"upper"`, `"title"`, atau `"sentence"`. |
| `Prefix` | `String` | `""` | Teks awalan sebelum hasil. |
| `Suffix` | `String` | `""` | Teks akhiran setelah satuan mata uang/sen. |
| `NegativePrefix` | `String` | `"minus"` | Kata pengganti tanda `-`. |
| `CurrencyUnit` | `String` | `"rupiah"` | Nama satuan mata uang utama. |
| `IncludeSen` | `Boolean` | `True` | Menyertakan satuan sen bila pecahan > `0`. |
| `SenUnit` | `String` | `"sen"` | Nama satuan pecahan sen. |
| `Rounding` | `String` | `"half-up"` | Mode pembulatan: `"half-up"` atau `"truncate"`. |

### 3. `FormatRupiah(Value, [Symbol], [SpaceAfterSymbol], [ThousandsSeparator], [DecimalSeparator], [Decimals], [Rounding], [NegativeFormat], [Compact], [CompactDecimals])`

| Parameter | Tipe | Default | Keterangan |
| :--- | :--- | :--- | :--- |
| `Value` | `Variant` | *(wajib)* | Angka, `Range` sel tunggal, atau teks numerik. |
| `Symbol` | `String` | `"Rp"` | Simbol mata uang (isi `""` jika tanpa simbol). |
| `SpaceAfterSymbol` | `Boolean` | `True` | Menambah spasi antara simbol dan angka. |
| `ThousandsSeparator` | `String` | `"."` | Pemisah ribuan. |
| `DecimalSeparator` | `String` | `","` | Pemisah desimal. |
| `Decimals` | `Variant` | `2` | Integer `0`–`20`, atau `"auto"` untuk menampilkan pecahan non-nol saja. |
| `Rounding` | `String` | `"half-up"` | Mode pembulatan: `"half-up"` atau `"truncate"`. |
| `NegativeFormat` | `String` | `"minus"` | `"minus"` (`"-Rp 1.000,00"`) atau `"parentheses"` (`"(Rp 1.000,00)"`). |
| `Compact` | `Boolean` | `False` | `True` untuk notasi ringkas (`rb`, `jt`, `M`, `T`). |
| `CompactDecimals` | `Variant` | `1` | Maksimal digit desimal saat `Compact = True` (`0`–`6`). |

### 4. `ParseRupiah(Value)`

Mengekstrak teks berformat Rupiah (`"Rp 1.500.000,50"`, `"IDR 2.500"`, `"(Rp 500.000)"`, `"-Rp 500.000"`) menjadi string numerik kanonik (`"1500000.50"`, `"-500000"`).

---

## Pengujian

### 1. Pengujian Native di dalam Excel VBA

1. Impor [`src/IndoTerbilang.bas`](./src/IndoTerbilang.bas) dan [`tests/IndoTerbilangTests.bas`](./tests/IndoTerbilangTests.bas) ke VBA Editor (`Alt + F11`).
2. Jalankan macro `RunAllIndoTerbilangTests` (`F5` atau `Alt + F8`).

### 2. Verifikasi Otomatis via Command Line (Windows `cscript.exe`)

Di Windows, Anda dapat menjalankan verifikasi otomatis terhadap seluruh logika [`src/IndoTerbilang.bas`](./src/IndoTerbilang.bas) melalui engine VBScript (`cscript.exe`):

```powershell
python tests/verify_macro.py
```

---

## Kredit & Referensi

Kaidah penulisan bilangan merujuk pada aturan baku **EYD V / KBBI** (Badan Pengembangan dan Pembinaan Bahasa).

## Lisensi

[MIT](./LICENSE) © 2026 salmankalista

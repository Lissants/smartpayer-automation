# SmartPayer Automation

Complete user manual for SmartPayer Automation, from first-time setup to daily operations.

Panduan lengkap SmartPayer Automation, mulai dari instalasi pertama hingga operasional harian.

## Languages / Bahasa

- [English](#english)
- [Bahasa Indonesia](#bahasa-indonesia)

---

# English

## Table of Contents

1. [Getting Started](#getting-started)
2. [Input Files](#input-files)
3. [Daily Rate](#daily-rate)
4. [Calendar Picker](#calendar-picker)
5. [Template Defaults](#template-defaults)
6. [Email Recipients](#email-recipients)
7. [Running the Automation](#running-the-automation)
8. [Managing Output](#managing-output)
9. [Retry and Troubleshooting](#retry-and-troubleshooting)

## Getting Started

### Launching SmartPayer

After the first-time setup, launch SmartPayer as follows:

1. Open the **SmartPayer Automation** folder.
2. Double-click `run_gui.bat`.
3. Wait for the SmartPayer graphical interface to open.

> **Important:** Use `run_gui.bat` every time you want to open SmartPayer. You do not need to open Python or any source code manually.

### First-Time Setup

The prerequisites only need to be installed once:

1. Locate `install_prerequisites.bat` in the SmartPayer folder.
2. Double-click the file.
3. Wait while the installer downloads and configures Python and the required libraries.
4. Wait for the confirmation message.
5. Close the terminal, then launch SmartPayer using `run_gui.bat`.

> **Terminal closes immediately?** Right-click `install_prerequisites.bat` and select **Run as administrator**. Some systems require elevated permissions to install the prerequisites.

The installer configures:

- **Python runtime**, which powers the automation engine.
- **Required libraries** for PDF generation, email delivery, and file handling.

## Input Files

SmartPayer uses a raw SAP export file as its input. You can load it in either of the following ways.

### Method A: Browse for a File

1. Click **Browse** on the main interface.
2. Navigate to the SAP export file.
3. Select the file and click **Open**.
4. Confirm that the selected path appears in SmartPayer.

### Method B: Use the INPUT Folder

1. Open the SmartPayer Automation folder.
2. Open the `INPUT` subfolder.
3. Copy or move the raw SAP export file into the folder.
4. SmartPayer will detect the file when the automation runs.

> **One file per run:** If the `INPUT` folder contains multiple files, SmartPayer uses the most recent file. Remove older input files to avoid processing the wrong file.

## Daily Rate

The default daily rate is **0.022% per day**.

### Keep the Default Rate

When prompted, click **Yes - Keep Rate** to continue using `0.022%` per day.

### Change the Rate

1. Click **No - Change Rate** when prompted.
2. Enter the new daily rate as a decimal percentage.
3. Confirm the value.

For example, enter:

```text
0.025
```

This represents `0.025%` per day. Do not include the `%` symbol.

> **Current run only:** A changed rate applies only to the current run. SmartPayer restores the default rate of `0.022%` the next time it starts.

## Calendar Picker

Use the calendar picker to specify the dates used when generating invoice letters.

- **Minimum:** 3 dates
- **Maximum:** 4 dates

### Selecting Dates

1. Click **Pick date** to open the calendar.
2. Select a date.
3. Click **Select**.
4. Repeat until at least three dates have been selected.
5. Optionally select a fourth date.

> SmartPayer will not proceed with fewer than three dates. It also prevents you from selecting more than four dates.

## Template Defaults

Template defaults control the variable content printed on each generated invoice letter. The settings remain saved until they are edited again.

### Company Information

- **Company name:** The full legal name printed on the letter, for example `PT. Mencari Cinta Sejati`.
- **Company abbreviation:** A short code used in headers and references, for example `MCS`.

### Financial Parameters

- **Deposito interest rate:** The interest rate printed in the invoice letter body.
- **Deposit days:** The number of deposit days referenced in the letter.

### Team Members

- **Team member names:** Names shown as signatories or contact persons.
- **Team member emails:** Email addresses shown as contact information.

To edit the defaults:

1. Open **Edit Defaults**.
2. Update the required company, financial, or team fields.
3. Click **Save Defaults**.

Changes take effect on the next automation run.

## Email Recipients

The recipient list maps each client name to one or more email addresses. SmartPayer uses this mapping to send each generated PDF to the correct recipient.

### View the Recipient List

Click **Recipient List** to view:

- Client or payer name
- Primary **To** address
- **CC** address

### Import Recipients

Prepare the recipient data in an `.xlsx` file, then use either method:

1. Click **Import Recipients** on the main interface, or
2. Open **Recipient List** and click **Import XLSX**.

### Manage Recipients

| Button | Function |
|---|---|
| **Delete all** | Deletes the entire recipient list. This cannot be undone. Re-import the XLSX file if needed. |
| **Save** | Saves additions and edits made during the current session. |

### Clients with Multiple Bill-To Parties

Create a separate recipient entry for each bill-to party number. Append the bill-to number to the client name.

```text
CV. INDRA JAYA 3001421  ->  recipient-a@email.com
CV. INDRA JAYA 3100432  ->  recipient-b@email.com
SEJATI MANDIRI 4002100  ->  recipient-c@email.com
```

This prevents letters for different bill-to parties from being sent to the same or incorrect address.

## Running the Automation

### Auto-Send Email

The **Auto-send email** option is enabled by default.

| Setting | SmartPayer behavior |
|---|---|
| **Enabled** | Generates one PDF invoice letter per payer and automatically emails each PDF to the corresponding recipient. |
| **Disabled** | Generates the PDF letters without sending email. The files remain in `Generated_Letters`. |

### Processing Flow

When auto-send is enabled:

```text
Raw SAP file -> Processing -> PDF letters -> Sent to recipients
```

When auto-send is disabled:

```text
Raw SAP file -> Processing -> PDFs saved locally
```

### Start a Run

1. Open SmartPayer using `run_gui.bat`.
2. Confirm that the input file is loaded through **Browse** or the `INPUT` folder.
3. Verify the daily rate.
4. Confirm that three or four dates are selected.
5. Review the template defaults and recipient list.
6. Enable or disable **Auto-send email** as required.
7. Click **Run Automation**.
8. Monitor the progress log and do not close the application while it is running.

After completion, the generated PDFs are stored in `Generated_Letters`. If auto-send was enabled, check the log for delivery failures.

## Managing Output

All generated PDFs are stored in `Generated_Letters`.

### Organise PDF Output

Click **Organise PDF Output** to group generated PDF files into subfolders based on their filename prefix.

```text
Generated_Letters/
├── SmartPayer April 2026/
│   ├── SmartPayer April 2026 - CV. ABADI.pdf
│   └── SmartPayer April 2026 - PT. MAJU.pdf
└── SmartPayer May 2026/
    └── SmartPayer May 2026 - CV. ABADI.pdf
```

Folders are created automatically when needed. Run this function after each automation session or before archiving a month's letters.

### Compress Completed PDFs

> **Deprecated:** This feature has been replaced by **Organise PDF Output** and can be ignored.

## Retry and Troubleshooting

If SmartPayer cannot match a payer name to an entry in the recipient list, the email delivery is marked as failed. You do not need to repeat the entire automation.

### Retry Failed Emails

1. Check the run log for entries marked **Failed**.
2. Note the affected client names.
3. Open **Recipient List**.
4. Add or correct the client entries and their **To** and **CC** addresses.
5. Click **Save**.
6. Click **Retry Failed Emails**.

SmartPayer resends only the previously failed letters. It does not regenerate PDFs, resend successful deliveries, or reprocess the SAP file.

### Common Issues

| Issue | Likely cause | Resolution |
|---|---|---|
| Delivery failed for a client | Client is not in the recipient list | Add the client, save, and retry failed emails. |
| Wrong recipient receives a letter | Duplicate client name without a bill-to suffix | Create separate entries using the bill-to number, for example `CV. INDRA JAYA 3001421`. |
| No PDFs are generated | Fewer than three dates selected or no input file loaded | Verify the input file and select at least three dates. |
| GUI does not open | `run_gui.bat` is missing or prerequisites are not installed | Run `install_prerequisites.bat`, then try `run_gui.bat` again. |
| Installer fails | Insufficient permissions | Right-click `install_prerequisites.bat` and select **Run as administrator**. |

---

# Bahasa Indonesia

## Daftar Isi

1. [Mulai dari Awal](#mulai-dari-awal)
2. [File Input](#file-input)
3. [Rate Harian](#rate-harian)
4. [Pilih Tanggal](#pilih-tanggal)
5. [Template Default](#template-default)
6. [Penerima Email](#penerima-email)
7. [Menjalankan Otomasi](#menjalankan-otomasi)
8. [Mengelola Output](#mengelola-output)
9. [Retry dan Penanganan Kendala](#retry-dan-penanganan-kendala)

## Mulai dari Awal

### Membuka SmartPayer

Setelah instalasi pertama selesai, buka SmartPayer dengan langkah berikut:

1. Buka folder **SmartPayer Automation**.
2. Klik dua kali `run_gui.bat`.
3. Tunggu sampai tampilan grafis SmartPayer terbuka.

> **Penting:** Jalankan `run_gui.bat` setiap kali ingin menggunakan SmartPayer. Kamu tidak perlu membuka Python atau source code secara manual.

### Instalasi Pertama

Prerequisite hanya perlu diinstal satu kali:

1. Cari `install_prerequisites.bat` di folder SmartPayer.
2. Klik dua kali file tersebut.
3. Tunggu selama installer mengunduh dan mengonfigurasi Python beserta library yang dibutuhkan.
4. Tunggu sampai muncul pesan konfirmasi.
5. Tutup terminal, lalu buka SmartPayer menggunakan `run_gui.bat`.

> **Terminal langsung tertutup?** Klik kanan `install_prerequisites.bat`, lalu pilih **Run as administrator**. Beberapa komputer memerlukan hak akses administrator untuk menginstal prerequisite.

Installer akan mengonfigurasi:

- **Python runtime** sebagai mesin utama otomasi.
- **Library pendukung** untuk membuat PDF, mengirim email, dan mengelola file.

## File Input

SmartPayer menggunakan file ekspor SAP mentah sebagai input. File dapat dimuat dengan salah satu cara berikut.

### Cara A: Pilih File dengan Browse

1. Klik **Browse** pada tampilan utama.
2. Buka lokasi file ekspor SAP.
3. Pilih file, lalu klik **Open**.
4. Pastikan path file muncul di SmartPayer.

### Cara B: Gunakan Folder INPUT

1. Buka folder SmartPayer Automation.
2. Buka subfolder `INPUT`.
3. Salin atau pindahkan file ekspor SAP mentah ke folder tersebut.
4. SmartPayer akan mendeteksi file saat otomasi dijalankan.

> **Satu file per proses:** Jika folder `INPUT` berisi beberapa file, SmartPayer akan menggunakan file yang paling baru. Hapus file lama untuk menghindari pemrosesan file yang salah.

## Rate Harian

Rate harian default adalah **0,022% per hari**.

### Menggunakan Rate Default

Saat diminta, klik **Yes - Keep Rate** untuk menggunakan `0.022%` per hari.

### Mengubah Rate

1. Klik **No - Change Rate** saat diminta.
2. Masukkan rate harian baru dalam format desimal persentase.
3. Konfirmasikan nilainya.

Contoh:

```text
0.025
```

Nilai tersebut berarti `0,025%` per hari. Jangan masukkan simbol `%`.

> **Hanya untuk proses saat ini:** Rate yang diubah hanya berlaku pada proses yang sedang berjalan. Saat SmartPayer dibuka kembali, rate akan kembali ke `0,022%`.

## Pilih Tanggal

Gunakan calendar picker untuk menentukan tanggal yang digunakan saat membuat surat tagihan.

- **Minimum:** 3 tanggal
- **Maksimum:** 4 tanggal

### Memilih Tanggal

1. Klik **Pick date** untuk membuka kalender.
2. Pilih tanggal.
3. Klik **Select**.
4. Ulangi sampai minimal tiga tanggal dipilih.
5. Jika diperlukan, pilih tanggal keempat.

> SmartPayer tidak dapat melanjutkan jika tanggal yang dipilih kurang dari tiga. Sistem juga mencegah pemilihan lebih dari empat tanggal.

## Template Default

Template default mengatur konten variabel yang dicetak pada setiap surat tagihan. Pengaturan akan tetap tersimpan sampai diubah kembali.

### Informasi Perusahaan

- **Nama perusahaan:** Nama resmi lengkap yang dicetak pada surat, misalnya `PT. Mencari Cinta Sejati`.
- **Singkatan perusahaan:** Kode singkat untuk header dan referensi surat, misalnya `MCS`.

### Parameter Keuangan

- **Suku bunga deposito:** Nilai suku bunga yang dicetak di isi surat tagihan.
- **Jumlah hari deposit:** Jumlah hari deposit yang dicantumkan di surat.

### Anggota Tim

- **Nama anggota tim:** Nama penandatangan atau kontak person pada surat.
- **Email anggota tim:** Alamat email yang dicantumkan sebagai informasi kontak.

Untuk mengubah pengaturan:

1. Buka **Edit Defaults**.
2. Ubah informasi perusahaan, parameter keuangan, atau data anggota tim.
3. Klik **Save Defaults**.

Perubahan akan digunakan pada proses otomasi berikutnya.

## Penerima Email

Daftar penerima menghubungkan nama klien dengan satu atau beberapa alamat email. SmartPayer menggunakan daftar ini untuk mengirim setiap PDF ke penerima yang sesuai.

### Melihat Daftar Penerima

Klik **Recipient List** untuk melihat:

- Nama klien atau payer
- Alamat utama **To**
- Alamat **CC**

### Import Penerima

Siapkan data penerima dalam file `.xlsx`, lalu gunakan salah satu cara berikut:

1. Klik **Import Recipients** di tampilan utama, atau
2. Buka **Recipient List**, lalu klik **Import XLSX**.

### Mengelola Penerima

| Tombol | Fungsi |
|---|---|
| **Delete all** | Menghapus seluruh daftar penerima. Tindakan ini tidak dapat dibatalkan. Import ulang file XLSX jika diperlukan. |
| **Save** | Menyimpan penambahan dan perubahan pada sesi saat ini. |

### Klien dengan Beberapa Bill-To Party

Buat entri penerima terpisah untuk setiap nomor bill-to party. Tambahkan nomor bill-to setelah nama klien.

```text
CV. INDRA JAYA 3001421  ->  penerima-a@email.com
CV. INDRA JAYA 3100432  ->  penerima-b@email.com
SEJATI MANDIRI 4002100  ->  penerima-c@email.com
```

Pemisahan ini mencegah surat dari bill-to party yang berbeda dikirim ke alamat yang sama atau penerima yang salah.

## Menjalankan Otomasi

### Auto-Send Email

Opsi **Auto-send email** aktif secara default.

| Pengaturan | Perilaku SmartPayer |
|---|---|
| **Aktif** | Membuat satu surat tagihan PDF per payer dan mengirim setiap PDF ke penerima yang sesuai. |
| **Tidak aktif** | Hanya membuat PDF tanpa mengirim email. File disimpan di `Generated_Letters`. |

### Alur Proses

Jika auto-send aktif:

```text
File SAP mentah -> Diproses -> Surat PDF -> Dikirim ke penerima
```

Jika auto-send tidak aktif:

```text
File SAP mentah -> Diproses -> PDF disimpan secara lokal
```

### Memulai Proses

1. Buka SmartPayer menggunakan `run_gui.bat`.
2. Pastikan file input sudah dimuat melalui **Browse** atau folder `INPUT`.
3. Periksa rate harian.
4. Pastikan tiga atau empat tanggal sudah dipilih.
5. Periksa template default dan daftar penerima.
6. Aktifkan atau nonaktifkan **Auto-send email** sesuai kebutuhan.
7. Klik **Run Automation**.
8. Pantau progress log dan jangan tutup aplikasi selama proses berjalan.

Setelah selesai, PDF tersimpan di `Generated_Letters`. Jika auto-send aktif, periksa log untuk melihat apakah ada pengiriman yang gagal.

## Mengelola Output

Semua PDF yang dibuat disimpan di `Generated_Letters`.

### Organise PDF Output

Klik **Organise PDF Output** untuk mengelompokkan file PDF ke dalam subfolder berdasarkan awalan nama file.

```text
Generated_Letters/
├── SmartPayer April 2026/
│   ├── SmartPayer April 2026 - CV. ABADI.pdf
│   └── SmartPayer April 2026 - PT. MAJU.pdf
└── SmartPayer May 2026/
    └── SmartPayer May 2026 - CV. ABADI.pdf
```

Folder akan dibuat secara otomatis jika belum tersedia. Jalankan fungsi ini setelah setiap sesi otomasi atau sebelum mengarsipkan surat untuk satu bulan.

### Compress Completed PDFs

> **Sudah tidak digunakan:** Fitur ini telah digantikan oleh **Organise PDF Output** dan dapat diabaikan.

## Retry dan Penanganan Kendala

Jika SmartPayer tidak dapat mencocokkan nama payer dengan daftar penerima, pengiriman email akan ditandai gagal. Seluruh proses otomasi tidak perlu diulang.

### Retry Email yang Gagal

1. Periksa log proses dan cari entri berstatus **Failed** atau **Gagal**.
2. Catat nama klien yang terdampak.
3. Buka **Recipient List**.
4. Tambahkan atau perbaiki data klien beserta alamat **To** dan **CC**.
5. Klik **Save**.
6. Klik **Retry Failed Emails**.

SmartPayer hanya mengirim ulang surat yang sebelumnya gagal. Sistem tidak membuat PDF baru, tidak mengirim ulang email yang sudah berhasil, dan tidak memproses ulang file SAP.

### Kendala Umum

| Kendala | Kemungkinan penyebab | Solusi |
|---|---|---|
| Pengiriman gagal untuk klien tertentu | Klien belum ada di daftar penerima | Tambahkan klien, simpan, lalu retry email yang gagal. |
| Surat diterima oleh penerima yang salah | Nama klien duplikat tanpa sufiks bill-to | Buat entri terpisah menggunakan nomor bill-to, misalnya `CV. INDRA JAYA 3001421`. |
| Tidak ada PDF yang dibuat | Tanggal kurang dari tiga atau file input belum dimuat | Periksa file input dan pilih minimal tiga tanggal. |
| GUI tidak terbuka | `run_gui.bat` tidak ditemukan atau prerequisite belum diinstal | Jalankan `install_prerequisites.bat`, lalu coba `run_gui.bat` kembali. |
| Installer gagal | Hak akses tidak mencukupi | Klik kanan `install_prerequisites.bat`, lalu pilih **Run as administrator**. |

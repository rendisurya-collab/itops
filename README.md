# Jira Logwork Telegram Bot

Bot Telegram untuk mempermudah isi logwork harian langsung ke Jira, jalan lokal di laptop kamu.

## Fitur

- `/log` — isi logwork baru ke issue Jira (issue key, waktu, deskripsi, tanggal), dengan konfirmasi sebelum dikirim
- `/today` — rekap logwork hari ini
- `/week` — rekap logwork minggu ini (Senin - hari ini)
- `/edit` — edit worklog yang sudah ada (pilih dari daftar 14 hari terakhir)
- `/delete` — hapus worklog
- Reminder otomatis setiap hari di jam yang kamu atur
- Hanya merespons akun Telegram kamu sendiri (berdasarkan `TELEGRAM_USER_ID`)

## 1. Siapkan Bot Telegram

1. Chat ke [@BotFather](https://t.me/BotFather) di Telegram.
2. Kirim `/newbot`, ikuti instruksi (kasih nama bot).
3. BotFather akan kasih **token**, contoh: `123456789:AAExxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx`. Simpan ini.
4. Cari tahu Telegram User ID kamu: chat ke [@userinfobot](https://t.me/userinfobot), kirim `/start`, dia akan balas ID kamu (angka).

## 2. Siapkan API Token Jira

1. Buka https://id.atlassian.com/manage-profile/security/api-tokens
2. Klik **Create API token**, kasih nama bebas, lalu copy token yang muncul (hanya muncul sekali).
3. Catat juga:
   - Base URL Jira kamu, contoh: `https://namaperusahaan.atlassian.net`
   - Email akun Jira kamu

## 3. Install & Konfigurasi

```bash
cd jira-logwork-bot
python3 -m venv venv
source venv/bin/activate      # Windows: venv\Scripts\activate
pip install -r requirements.txt

cp .env.example .env
```

Buka file `.env`, isi semua nilainya:

```
TELEGRAM_BOT_TOKEN=...
TELEGRAM_USER_ID=...
JIRA_BASE_URL=https://namaperusahaan.atlassian.net
JIRA_EMAIL=nama@email.com
JIRA_API_TOKEN=...
REMINDER_HOUR=17
REMINDER_MINUTE=30
TIMEZONE=Asia/Jakarta
```

## 4. Jalankan Bot

```bash
python3 bot.py
```

Bot akan jalan selama terminal ini terbuka. Buka chat dengan bot kamu di Telegram, kirim `/start`.

Untuk menghentikan, tekan `Ctrl+C` di terminal.

## Cara Pakai

### Daftarkan akun Jira pribadi (wajib, kalau bot dipakai bareng-bareng di grup)

Kalau bot ini dipakai sendirian, akun Jira di `.env` (`JIRA_EMAIL`/`JIRA_API_TOKEN`) sudah cukup. Tapi kalau bot dipakai **bareng-bareng di grup** (banyak orang), tiap orang harus daftarkan akun Jira-nya sendiri dulu, supaya logwork/edit/delete tercatat atas nama masing-masing (bukan atas nama 1 akun yang sama):

```
/myjira
```
Bot akan minta:
1. Email akun Jira kamu
2. API token Jira kamu (kalau belum punya, bot kasih link untuk generate: https://id.atlassian.com/manage-profile/security/api-tokens)

Bot langsung verifikasi ke Jira, kalau berhasil, tinggal konfirmasi simpan. Setelah ini, `/log`, `/edit`, `/delete`, `/today`, `/week` otomatis pakai akun Jira kamu sendiri (bukan punya orang lain).

Cek status akun kamu: `/myjira status` &nbsp; | &nbsp; Hapus akun: `/myjira remove`

> `/tasks` dan `/export` tetap pakai 1 akun bersama (akun admin di `.env`) karena itu cuma fitur "lihat" data project, bukan nulis/logwork -- jadi tidak perlu akun personal.

**Cara cepat langsung eksekusi (skip lihat isi guidance dulu):**
```
/run delete reservation
```
Kalau ketemu 1 guidance yang cocok dan punya action, bot langsung lanjut ke step minta parameter (skip step tampilkan isi guidance & tombol). Kalau ketemu beberapa yang cocok, bot kasih tombol pilihan. Kalau guidance-nya belum punya action, bot kasih tahu.

### Guidance yang bisa dieksekusi otomatis (Action/Script)

Selain sekadar menampilkan panduan, guidance juga bisa dilengkapi **action** — script PowerShell (`.ps1`) yang langsung dijalankan bot di laptop kamu, lengkap dengan parameter yang kamu isi lewat chat.

**Contoh alur (sesuai request awal):**
```
kamu: delete reservation
bot: 📘 delete reservation
     (isi guidance...)
     [⚙️ Jalankan Script]   <- tombol
kamu: tap tombol itu
bot: Masukkan nilai untuk parameter -ReservationId
kamu: 12345, 12346
bot: Konfirmasi:
     Script: D:\promovoucher\cancel_reservation.ps1
     Parameter -ReservationId: 12345, 12346
     [✅ Ya, jalankan] [Batal]
kamu: tap "Ya, jalankan"
bot: ⏳ Menjalankan cancel_reservation.ps1...
bot: ✅ Selesai (exit code 0):
     (output/log dari script ditampilkan di sini)
```

**Cara menambahkan action ke guidance baru:**
Saat `/addguide`, setelah isi guidance, bot akan tanya "Apakah guidance ini bisa dieksekusi otomatis?". Kalau **Ya**, bot minta:
1. Path lengkap file script di laptop kamu, contoh: `D:\promovoucher\cancel_reservation.ps1`
2. Nama parameter/flag sesuai script. Bisa **1 parameter** (contoh: `-CsvFile`) atau **banyak parameter sekaligus**, dipisah koma (contoh: `-ArticleId, -SourceId`)
3. Kalau cuma 1 parameter, bot tanya bentuk input-nya: **teks bebas** atau **upload file**. Kalau lebih dari 1 parameter, otomatis jadi teks bebas semua (upload file cuma didukung untuk 1 parameter).

Waktu dijalankan nanti, bot akan tanya nilai tiap parameter **satu per satu** secara berurutan. Kalau satu parameter butuh lebih dari 1 nilai, tinggal pisahkan dengan koma saat menjawab, contoh:
```
Masukkan nilai untuk parameter -ArticleId
(kalau lebih dari 1 nilai, pisahkan dengan koma)
> 8000044321, 8000044322
```
Ini otomatis diteruskan ke script sebagai array (`-ArticleId "8000044321" "8000044322"`), sesuai cara PowerShell menerima parameter array.

**Mode input lain: baris per baris (berpasangan)** — kalau ada lebih dari 1 parameter dan nilai-nilainya berpasangan (misal tiap Article ID punya Source ID sendiri), pas setup action pilih **"Baris per baris (berpasangan)"**. Nanti waktu dijalankan, bot cuma tanya sekali dan kamu isi semua baris sekaligus:
```
Masukkan data, satu baris per kombinasi (ArticleId, SourceId):

> 8000044321, SS20
> 8000044322, SS21
```
Baris pertama otomatis jadi `-ArticleId 8000044321 -SourceId SS20`, baris kedua `-ArticleId 8000044322 -SourceId SS21`, digabung jadi array seperti mode biasa.

**Cara menambahkan/menghapus action ke guidance yang sudah ada:** pakai `/editguide`, lalu pilih tombol **"⚙️ Tambah Action"** / **"⚙️ Hapus Action"**.

> ⚠️ **Penting soal keamanan:** siapa pun yang bisa akses akun Telegram kamu (dan `TELEGRAM_USER_ID` di `.env` cocok) bisa menjalankan script ini lewat bot. Karena itu:
> - Selalu ada langkah **konfirmasi** sebelum script benar-benar dijalankan
> - Jangan pernah share token bot / akses akun Telegram kamu ke orang lain
> - Untuk script yang sifatnya **merusak/menghapus data** (seperti contoh cancel/delete), pastikan script itu sendiri sudah aman (misal ada validasi input, logging, atau dry-run mode) sebelum dihubungkan ke bot
> - Output/log script (sampai ±3500 karakter) akan dikirim balik ke chat sebagai bukti hasil eksekusi

### Guidance / panduan tersimpan

Simpan panduan-panduan yang sering kamu pakai (SOP, cara troubleshoot, dsb) supaya gampang dipanggil lagi lewat bot.

**Tambah guidance baru:**
```
/addguide
> Judul, contoh: Cara mengatasi koneksi jaringan error
> Kata kunci pemicu (pisah koma), contoh: koneksi jaringan error, jaringan error, network issue
> Isi/detail guidance-nya (boleh panjang, multi-baris)
> Konfirmasi simpan
```

> Bisa juga upload **file** (script, dokumen, dll) di step "isi guidance" — tinggal kirim file-nya langsung, dan caption di file tersebut akan dipakai sebagai deskripsinya. File ini akan ikut dikirim ulang setiap kali guidance-nya dipanggil.

**Panggil guidance:**
```
/guide koneksi jaringan error
```
atau langsung ketik kalimatnya tanpa command sama sekali, bot otomatis kenali:
```
cara mengatasi koneksi jaringan error
```
Kalau ada beberapa guidance yang cocok, bot akan kasih tombol pilihan.

**Lihat semua guidance tersimpan:**
```
/listguide
```

**Edit guidance:**
```
/editguide
```
Pilih guidance dari daftar, lalu pilih bagian yang mau diubah (judul / kata kunci / isi-file), bisa ubah beberapa bagian sekaligus sebelum tekan "Selesai & Simpan".

**Hapus guidance:**
```
/delguide
```
(pilih dari daftar yang muncul)

**Edit manual lewat file:** semua guidance disimpan di file `guidance.json` (otomatis dibuat di folder yang sama dengan `bot.py`). Kamu bisa edit file ini langsung pakai text editor / Notepad kalau mau — bot selalu baca ulang file ini setiap kali dibutuhkan, jadi perubahan langsung kepakai tanpa perlu restart bot. Formatnya:
```json
[
  {
    "id": "0001",
    "title": "Cara mengatasi koneksi jaringan error",
    "keywords": ["koneksi jaringan error", "jaringan error", "network issue"],
    "content": "1. Cek kabel LAN/WiFi\n2. Restart router\n3. Restart laptop\n4. Hubungi IT kalau masih error"
  }
]
```

### Download report Excel

```
/export TIC
```
Bot akan generate file Excel berisi daftar task di project itu, dengan kolom: **No, Task SD, Task Dev, Summary, Status Dev, Status QA, Created, Tanggal Dev Done, Tanggal QA Done**.

Cara bot menentukan **Task Dev** dan **Status QA**: bot melihat *linked issues* bertipe **"relates to"** pada tiap task. Di antara issue yang ter-link:
- yang judulnya **diawali kata "Testing"** dianggap task QA → statusnya masuk kolom **Status QA**
- issue lain (non-Testing) dianggap **Task Dev** → key & statusnya masuk kolom **Task Dev** / **Status Dev**

**Tanggal Dev Done** dan **Tanggal QA Done** diambil dari riwayat (history) perubahan status task Dev/QA tersebut — tanggal terakhir kali statusnya berubah jadi "Done". Kalau task itu belum pernah "Done", kolomnya dikosongkan.

Kalau suatu task tidak punya linked issue, kolom-kolom terkait dikosongkan saja.

> Catatan: karena bot perlu cek riwayat status tiap task Dev & QA satu per satu ke Jira, proses `/export` untuk project dengan banyak task bisa memakan waktu agak lama (beberapa detik sampai beberapa menit, tergantung jumlah task).

### Lihat daftar task di suatu project

Kalau kamu tidak hafal semua kode task, tinggal minta bot tampilkan daftarnya:
```
/tasks TDBU
```
atau lebih cepat, cukup ketik kode project-nya langsung (tanpa `/tasks`):
```
TDBU-
```
```
TIC
```
Bot akan balas daftar 50 task terbaru (diurutkan dari yang terakhir diupdate) di project itu, lengkap dengan judul dan status masing-masing.

> Catatan: kalau tanpa `/tasks`, kode project harus **HURUF BESAR semua** (misal `TIC`, bukan `tic`) supaya tidak salah kena kata-kata biasa. Kalau pakai tanda `-` di belakang (misal `TDBU-`), huruf besar/kecil bebas.

### Isi logwork baru

**Cara 1 - langkah demi langkah:**
```
/log
> masukkan issue key, contoh: PROJ-123
> masukkan waktu, contoh: 2h 30m
> masukkan deskripsi
> pilih tanggal (Hari ini / Kemarin / manual)
> konfirmasi kirim
```

**Cara 2 - satu kalimat langsung (bahasa natural):**
```
/log saya mengerjakan aktifitas di TDBU-10 pada jam 10 pagi, deskripsi pekerjaanya analyze issue price
```
Bot otomatis mengenali:
- **Issue key**: pola seperti `TDBU-10`, `PROJ-123`
- **Jam**: frasa `jam <angka> pagi/siang/sore/malam` (contoh: `jam 10 pagi`, `jam 2 siang`, `jam 8 malam`)
- **Deskripsi**: teks setelah kata `deskripsi` / `deskripsi pekerjaannya`
- **Durasi** (opsional): frasa `selama ...` atau `durasi ...`, contoh: `selama 2 jam 30 menit`
- **Kemarin** (opsional): kalau kalimat mengandung kata `kemarin`, tanggalnya otomatis jadi kemarin

Field apa pun yang **tidak disebutkan** di kalimat (misalnya durasi), akan **ditanyakan otomatis** oleh bot setelahnya. Jadi kalimat di atas tetap akan lanjut bertanya "Berapa lama waktu yang dihabiskan?" karena durasinya belum disebut.

Contoh lain yang mengisi semua sekaligus (langsung ke konfirmasi tanpa tanya apa-apa lagi):
```
/log TDBU-10 selama 2 jam jam 10 pagi deskripsi analyze issue price
```


### Lihat rekap
```
/today   -> rekap hari ini
/week    -> rekap minggu berjalan
```

### Edit / hapus logwork
```
/edit    -> pilih issue, lalu pilih worklog dari daftar, lalu isi perubahan
/delete  -> pilih issue, lalu pilih worklog, lalu konfirmasi hapus
```

Ketik `/cancel` kapan saja untuk membatalkan proses yang sedang berjalan.

## Kirim notifikasi ke grup Telegram (opsional)

Semua notifikasi/reminder otomatis (reminder logwork, tiket ServiceDesk Plus baru, reminder tiket Open) bisa diarahkan ke **grup** Telegram, bukan cuma ke chat personal kamu.

**Cara setup:**
1. Invite bot kamu ke grup Telegram yang dituju
2. Di dalam grup itu, ketik:
   ```
   /chatid
   ```
   Bot akan balas ID grup-nya, contoh: `-1001234567890` (grup selalu diawali tanda minus `-`)
3. Copy ID itu, isi di `.env`:
   ```
   TELEGRAM_GROUP_ID=-1001234567890
   ```
4. Restart bot

Setelah ini, reminder logwork, notif tiket baru SDP, dan reminder tiket Open akan otomatis masuk ke grup tersebut. Balasan command yang kamu ketik langsung (seperti `/log`, `/sdtickets`, dll) tetap dibalas di chat tempat kamu mengetiknya (personal atau grup, sesuai kamu ngetik di mana).

Kosongkan `TELEGRAM_GROUP_ID` (atau hapus barisnya) untuk balik ke notifikasi personal seperti semula.

**Kalau grup kamu pakai mode Topics/Forum** (dan kamu mau notifikasi masuk ke topic tertentu saja, bukan percakapan umum grup):
1. Buka topic yang dituju (contoh: topic khusus bot)
2. Di dalam topic itu, ketik `/chatid` -- bot akan balas ID grup **dan** ID topic-nya
3. Isi di `.env`:
   ```
   TELEGRAM_TOPIC_ID=6
   ```
4. Restart bot

Semua notifikasi otomatis akan masuk ke topic itu secara spesifik. Kosongkan `TELEGRAM_TOPIC_ID` untuk kirim ke percakapan umum grup (bukan topic tertentu).

### Kirim ke banyak grup/topic sekaligus

Kalau kamu mau notifikasi masuk ke **lebih dari 1 grup atau topic** sekaligus (misal ke topic tim Dev L1, L2, L3 yang beda-beda, atau ke beberapa grup berbeda), pakai `TELEGRAM_NOTIFY_TARGETS` di `.env` (isi ini SAJA, tidak usah pakai `TELEGRAM_GROUP_ID`/`TELEGRAM_TOPIC_ID` lagi):

```
TELEGRAM_NOTIFY_TARGETS=-1001111111111:6,-1002222222222,-1003333333333:3
```

Formatnya `ChatID:TopicID` dipisah koma:
- `-1001111111111:6` → grup ini, topic nomor 6
- `-1002222222222` → grup ini, tanpa topic (grup biasa / percakapan umum)
- `-1003333333333:3` → grup lain, topic nomor 3

Cara dapatkan tiap ChatID & TopicID-nya sama seperti sebelumnya: masuk ke grup/topic yang dituju, ketik `/chatid` di situ.

Setelah `TELEGRAM_NOTIFY_TARGETS` diisi dan bot di-restart, **semua** target itu akan menerima notifikasi yang sama (reminder logwork, notif tiket SDP, dll) sekaligus, dan **semua anggota** dari tiap grup yang terdaftar otomatis boleh pakai command bot juga (lihat catatan multi-user di bawah).

> **Penting -- ini juga membuka akses multi-user:** begitu `TELEGRAM_GROUP_ID` diisi, **semua anggota grup itu otomatis boleh pakai semua command bot** (termasuk `/run` yang bisa eksekusi script, `/delete` worklog, dll), bukan cuma kamu. Kalau bot ini dipakai bareng-bareng oleh tim (misal 6 anggota), inilah cara mengaktifkannya. Kalau kamu **tidak** mau anggota grup bisa pakai bot (cuma mau notifikasi masuk ke grup saja, command tetap personal), jangan pakai `TELEGRAM_GROUP_ID` -- kasih tahu saya, ada cara lain yang bisa dipisah antara "notify ke grup" dan "siapa yang boleh command".
>
> Supaya worklog Jira dari tiap anggota grup tercatat atas nama masing-masing (bukan 1 akun yang sama), tiap orang wajib daftarkan akun Jira pribadinya dulu lewat `/myjira` (lihat bagian "Daftarkan akun Jira pribadi" di atas). Anggota yang belum daftar akan diminta `/myjira` dulu setiap kali coba pakai `/log`, `/edit`, `/delete`, `/today`, atau `/week`.

## Menghubungkan ke ServiceDesk Plus (opsional)

Selain Jira, bot ini juga bisa dihubungkan ke **ManageEngine ServiceDesk Plus** (misal untuk narik tiket/laporan dari sana).

**1. Generate API Key:**
1. Login ke ServiceDesk Plus kamu
2. Klik ikon profil (pojok kanan atas) → **Generate API Key**
3. Copy key yang muncul

**2. Isi di `.env`:**
```
SDP_BASE_URL=https://servicedesk.namaperusahaan.com
SDP_API_KEY=isi_api_key_kamu
```

**3. Restart bot**, lalu coba:
```
/sdtickets
```
> Kalau `SDP_NOTIFY_GROUPS` sudah diisi di `.env`, `/sdtickets` **otomatis kebatasi** ke group-group itu saja (sama seperti fitur notifikasi & reminder). Kalau belum diisi, tampil semua group.

Filter tambahan berdasarkan status dan/atau group (urutan bebas, bisa salah satu atau dua-duanya):
```
/sdtickets status:Open
/sdtickets group:Network Support
/sdtickets status:Open group:Network Support
```
Mau lihat **semua group** (bypass default `SDP_NOTIFY_GROUPS`):
```
/sdtickets group:all
```
> Nama group harus diketik **persis sama** dengan yang ada di ServiceDesk Plus. Cek nama-nama group-nya langsung dari web ServiceDesk Plus (biasanya di **Admin → Groups**, atau lihat kolom "Group" di halaman daftar tiket) — ServiceDesk Plus tidak menyediakan API untuk menarik daftar group secara otomatis, jadi ini perlu dicek manual.

Lihat detail 1 tiket:
```
/sdticket 12345
```

> **Catatan:** karena tiap instalasi ServiceDesk Plus bisa beda konfigurasi/versinya, kemungkinan besar perlu ada penyesuaian kecil setelah test pertama kali (mirip waktu setup Jira dulu). Kalau ada error, screenshot aja hasilnya untuk didebug bareng.

### Notifikasi tiket baru masuk (berdasarkan group)

Bot bisa otomatis cek tiket baru secara berkala, dan kirim notif kalau ada yang masuk di group yang kamu pantau.

**Setup di `.env`:**
```
SDP_NOTIFY_GROUPS=IT Digital - TechDev Ops L1,IT Digital - TechDev Ops L2,IT Digital - TechDev Ops L3
SDP_NOTIFY_INTERVAL_MINUTES=5
```
- `SDP_NOTIFY_GROUPS`: nama-nama group yang dipantau, dipisah koma, harus **persis sama** dengan nama di ServiceDesk Plus. Kosongkan untuk menonaktifkan fitur ini.
- `SDP_NOTIFY_INTERVAL_MINUTES`: bot cek tiket baru tiap berapa menit (default 5 menit).

Setelah restart bot, notifikasi otomatis aktif — tidak perlu command apa pun. Bentuk notifnya:
```
🆕 Tiket baru masuk (IT Digital - TechDev Ops L2):

• [12345] Server down, tidak bisa akses
    status: Open | requester: Budi Santoso
```

**Cara kerjanya:** bot menyimpan ID tiket terakhir yang sudah dicek di file `sdp_notify_state.json` (otomatis dibuat di folder yang sama dengan `bot.py`). Pertama kali bot dijalankan, tidak ada notif yang dikirim (cuma menyimpan baseline supaya tidak spam notif tiket-tiket lama) — notif baru mulai muncul untuk tiket yang masuk **setelah** bot pertama kali jalan.

> Fitur ini hanya jalan selama bot sedang berjalan di laptop kamu (`python bot.py` aktif).

### Reminder berkala untuk tiket yang masih Open

Selain notif sekali waktu tiket baru masuk, bot juga bisa **terus mengingatkan** selama tiket berstatus Open — intervalnya bisa kamu ubah kapan saja langsung dari chat, tanpa perlu restart bot.

```
/sdreminder 30        -> aktifkan, cek & ingatkan tiap 30 menit
/sdreminder 15        -> ganti jadi tiap 15 menit
/sdreminder off        -> matikan
/sdreminder status     -> cek status & interval saat ini
/sdreminder            -> sama seperti status
```

Bot akan kirim daftar semua tiket yang masih Open (di group yang sama dengan `SDP_NOTIFY_GROUPS`) tiap kali intervalnya tercapai. Kalau tidak ada tiket Open sama sekali, bot tidak kirim apa-apa (tidak spam).

Bisa juga diatur nilai awalnya lewat `.env`:
```
SDP_OPEN_REMINDER_MINUTES=30
```
Tapi ini cuma nilai default kalau belum pernah diatur lewat `/sdreminder` — begitu kamu ubah lewat chat, itu yang dipakai (tersimpan di file `sdp_notify_state.json`, tetap kepakai walau bot di-restart).

## Catatan

- Format waktu logwork mengikuti format Jira: `2h`, `1h 30m`, `45m`, `1d`, dst.
- `/edit` dan `/delete` hanya menampilkan worklog milik kamu sendiri (dicocokkan lewat email) dalam 14 hari terakhir.
- Reminder harian dikirim otomatis sesuai jam di `.env`, selama bot sedang berjalan di jam tersebut.
- Bot ini jalan lokal — kalau laptop mati/tidur atau bot tidak dijalankan, reminder dan bot tidak akan aktif. Kalau nanti mau reminder selalu jalan, bot ini bisa dipindah ke server kecil (VPS) atau layanan seperti Railway/Render.

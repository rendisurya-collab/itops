import time
import schedule
from bot_notifier import run_adminer_playwright

def job():
    print(f"\n[SCHEDULER] Menjalankan pengecekan bot pada: {time.strftime('%Y-%m-%d %H:%M:%S')}")
    try:
        run_adminer_playwright()
    except Exception as e:
        print(f"[SCHEDULER ERROR] Gagal mengeksekusi bot: {e}")

# --- PENGATURAN JADWAL --- 
# Jalankan setiap 15 menit
# --- PENGATURAN JADWAL ---

# Contoh 1: Setiap N Menit (ganti angka 15 sesuai keinginan)
#schedule.every(10).minutes.do(job)

# Per jam
schedule.every(3).hour.do(job)

# Per hari jam 09:00
#schedule.every().day.at("09:00").do(job)

# Opsi lain (pilih/uncomment salah satu jika dibutuhkan):
# schedule.every(1).hours.do(job)              # Setiap 1 jam
# schedule.every().day.at("08:00").do(job)      # Setiap hari jam 08:00 pagi

print("[INFO] Scheduler berjalan! Bot akan mengeksekusi query secara berkala...")
print("[INFO] Tekan Ctrl+C untuk menghentikan scheduler.\n")

# Jalankan eksekusi pertama kali secara langsung saat scheduler diaktifkan
job()

# Loop utama untuk memantau jadwal
while True:
    schedule.run_pending()
    time.sleep(1)
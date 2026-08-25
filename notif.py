import asyncio
from desktop_notifier import DesktopNotifier

notifier = DesktopNotifier()

async def main():
    # Kirim notifikasi sederhana
    await notifier.send(
        title="Halo dari Python!",
        message="Script saya sudah selesai dijalankan.",
    )
    print("Notifikasi telah dikirim!")

# Jalankan fungsi async
asyncio.run(main())
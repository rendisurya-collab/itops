import os
from dotenv import load_dotenv

load_dotenv()

TELEGRAM_BOT_TOKEN = os.getenv("TELEGRAM_BOT_TOKEN", "")
TELEGRAM_USER_ID = os.getenv("TELEGRAM_USER_ID", "")

# ID grup Telegram (opsional, format lama -- 1 tujuan saja). Kalau diisi,
# notifikasi otomatis dikirim ke grup ini. Cara dapatkan: invite bot ke grup,
# ketik /chatid di grup itu.
TELEGRAM_GROUP_ID = os.getenv("TELEGRAM_GROUP_ID", "").strip()

# ID Topic di dalam grup (opsional, cuma berlaku kalau grup-nya mode Forum/Topics
# aktif, dipakai bareng TELEGRAM_GROUP_ID di atas). Cara dapatkan: buka topic-nya,
# ketik /chatid di dalam topic itu -- bot balas ID grup DAN ID topic-nya.
TELEGRAM_TOPIC_ID = os.getenv("TELEGRAM_TOPIC_ID", "").strip()

# Banyak tujuan sekaligus (opsional, format baru). Isi salah satu ini ATAU
# TELEGRAM_GROUP_ID/TELEGRAM_TOPIC_ID di atas -- kalau ini diisi, yang di atas
# diabaikan. Format: "ChatID:TopicID" dipisah koma. TopicID boleh dikosongkan
# kalau grup itu tidak pakai Topics (langsung tulis ChatID-nya saja).
# Contoh: -1001111111111:6,-1002222222222,-1003333333333:3
TELEGRAM_NOTIFY_TARGETS_RAW = os.getenv("TELEGRAM_NOTIFY_TARGETS", "").strip()


def _parse_notify_targets():
    targets = []
    if TELEGRAM_NOTIFY_TARGETS_RAW:
        for item in TELEGRAM_NOTIFY_TARGETS_RAW.split(","):
            item = item.strip()
            if not item:
                continue
            if ":" in item:
                chat_part, topic_part = item.split(":", 1)
                chat_part = chat_part.strip()
                topic_part = topic_part.strip()
                thread_id = int(topic_part) if topic_part.isdigit() else None
            else:
                chat_part, thread_id = item, None
            if chat_part:
                targets.append((chat_part, thread_id))
    elif TELEGRAM_GROUP_ID:
        thread_id = int(TELEGRAM_TOPIC_ID) if TELEGRAM_TOPIC_ID.isdigit() else None
        targets.append((TELEGRAM_GROUP_ID, thread_id))
    elif TELEGRAM_USER_ID:
        targets.append((TELEGRAM_USER_ID, None))
    return targets


NOTIFY_TARGETS = _parse_notify_targets()


def notify_targets():
    """List of (chat_id, message_thread_id) tujuan pengiriman notifikasi/reminder otomatis."""
    return NOTIFY_TARGETS


def allowed_chat_ids():
    """Semua chat_id yang boleh pakai command bot: personal owner + semua grup
    yang terdaftar di notify_targets()."""
    ids = {str(TELEGRAM_USER_ID)}
    for chat_id, _ in NOTIFY_TARGETS:
        if chat_id:
            ids.add(str(chat_id))
    return ids

JIRA_BASE_URL = os.getenv("JIRA_BASE_URL", "").rstrip("/")
JIRA_EMAIL = os.getenv("JIRA_EMAIL", "")
JIRA_API_TOKEN = os.getenv("JIRA_API_TOKEN", "")

REMINDER_HOUR = int(os.getenv("REMINDER_HOUR", "17"))
REMINDER_MINUTE = int(os.getenv("REMINDER_MINUTE", "30"))

def validate():
    missing = []
    for key in [
        "TELEGRAM_BOT_TOKEN",
        "TELEGRAM_USER_ID",
        "JIRA_BASE_URL",
        "JIRA_EMAIL",
        "JIRA_API_TOKEN",
    ]:
        if not globals().get(key):
            missing.append(key)
    if missing:
        raise RuntimeError(
            f"Konfigurasi belum lengkap di file .env, isi dulu: {', '.join(missing)}"
        )

    # Log peringatan jika ServiceDesk Plus belum aktif
    if not sdp_configured():
        print("⚠️ WARNING: SDP_BASE_URL atau SDP_API_KEY belum diisi di .env. Otomasi SDP akan nonaktif.")
    elif not SDP_NOTIFY_GROUPS:
        print("⚠️ WARNING: SDP_NOTIFY_GROUPS masih kosong di .env. Otomasi SDP akan nonaktif.")

# Kalau REMINDER_INTERVAL_MINUTES diisi (>0), reminder akan berulang tiap
# sekian menit (contoh: 30) selama jam kerja, dan REMINDER_HOUR/MINUTE di atas
# diabaikan. Kosongkan / isi 0 untuk pakai reminder sekali sehari seperti biasa.
REMINDER_INTERVAL_MINUTES = int(os.getenv("REMINDER_INTERVAL_MINUTES", "0"))
REMINDER_START_HOUR = int(os.getenv("REMINDER_START_HOUR", "9"))
REMINDER_END_HOUR = int(os.getenv("REMINDER_END_HOUR", "18"))

TIMEZONE = os.getenv("TIMEZONE", "Asia/Jakarta")

# ==== ServiceDesk Plus (opsional) ====
SDP_BASE_URL = os.getenv("SDP_BASE_URL", "").rstrip("/")
SDP_API_KEY = os.getenv("SDP_API_KEY", "")

# Notifikasi tiket baru (opsional). Isi nama-nama group dipisah koma untuk
# mengaktifkan; kosongkan untuk menonaktifkan.
SDP_NOTIFY_GROUPS = [g.strip() for g in os.getenv("SDP_NOTIFY_GROUPS", "").split(",") if g.strip()]
SDP_NOTIFY_INTERVAL_MINUTES = int(os.getenv("SDP_NOTIFY_INTERVAL_MINUTES", "5"))

# Reminder berkala untuk tiket yang masih berstatus Open. Nilai default kalau
# belum pernah diatur lewat command /sdreminder. 0 = nonaktif.
SDP_OPEN_REMINDER_DEFAULT_MINUTES = int(os.getenv("SDP_OPEN_REMINDER_MINUTES", "0"))

# Monitor SLA/OverDue: kirim notif kalau ada tiket yang sudah lewat DueBy (SLA).
# SDP_SLA_MONITOR_ENABLED: aktif/tidak. SDP_SLA_CHECK_INTERVAL_MINUTES: interval cek.
SDP_SLA_MONITOR_ENABLED = os.getenv("SDP_SLA_MONITOR_ENABLED", "true").strip().lower() in ("1", "true", "yes")
SDP_SLA_CHECK_INTERVAL_MINUTES = int(os.getenv("SDP_SLA_CHECK_INTERVAL_MINUTES", "30"))


def sdp_configured() -> bool:
    return bool(SDP_BASE_URL and SDP_API_KEY)


def validate():
    missing = []
    for key in [
        "TELEGRAM_BOT_TOKEN",
        "TELEGRAM_USER_ID",
        "JIRA_BASE_URL",
        "JIRA_EMAIL",
        "JIRA_API_TOKEN",
    ]:
        if not globals().get(key):
            missing.append(key)
    if missing:
        raise RuntimeError(
            f"Konfigurasi belum lengkap di file .env, isi dulu: {', '.join(missing)}"
        )

# Pemetaan "Nama Department": "Nama Subcategory Target"
DEPARTMENT_SUBCATEGORY_MAP = {
    "SS After Sales - SS After Sales": "Repair",
    "SS E-Commerce Operation - CHANNEL OPERATION DEPT": "Eraspace",
    "ED CM Monobrand & Online - CM ONLINE DEPT": "Eraspace",
    "ED CM B2B - CM B2B PORTAL OPERATION & DEVELOPMENT DEPT": "Dealer Portal",
    "ED Finance Operation - ED AR RETAIL & E-COMMERCE DEPT": "Dealer Portal",
    "SS Product & Digital Innovation - SS Digital Product Management": "Eraspace",
    "EAL JDS Multichannel - EAL JDS CHANNEL OPERATION DEPT": "JD Sports"
}

# Fallback subcategory jika tiket tidak punya subcategory dan department tidak cocok mapping
SDP_DEFAULT_SUBCATEGORY = os.getenv("SDP_DEFAULT_SUBCATEGORY", "General")

# ==== SLA Monitor (Playwright scraping dashboard SDP) ====
SLA_MONITOR_ENABLED = os.getenv("SLA_MONITOR_ENABLED", "true").strip().lower() in ("1", "true", "yes")
SLA_MONITOR_INTERVAL_MINUTES = int(os.getenv("SLA_MONITOR_INTERVAL_MINUTES", "180"))
SLA_MONITOR_URL = os.getenv("SLA_MONITOR_URL", "https://servicedesk.erajaya.com/WOListView.do")
SLA_MONITOR_STATE_FILE = os.getenv("SLA_MONITOR_STATE_FILE", "state.json")

# ==== Google Sheets Logging ====
GOOGLE_SHEETS_CREDENTIALS = os.getenv("GOOGLE_SHEETS_CREDENTIALS", "")
GOOGLE_SHEETS_SPREADSHEET_ID = os.getenv("GOOGLE_SHEETS_SPREADSHEET_ID", "18Vku1ikHbF8sa4ioRMm-yfkSejAUH4nsr7vE5JcjGI8")

# ==== Web Search (Serper.dev) ====
# Toggle on/off General Query Handler (web search) di /tanyabot.
# Set "false" untuk menonaktifkan tanpa perlu menghapus API key.
WEB_SEARCH_ENABLED = os.getenv("WEB_SEARCH_ENABLED", "true").strip().lower() in ("1", "true", "yes")
# API key dari https://serper.dev (gratis 2500 query). Kalau kosong, fitur web
# search di /tanyabot dinonaktifkan (bot langsung fallback simpan pending).
SERPER_API_KEY = os.getenv("SERPER_API_KEY", "")
# Region & bahasa hasil pencarian (default Indonesia)
SERPER_GL = os.getenv("SERPER_GL", "id")
SERPER_HL = os.getenv("SERPER_HL", "id")

# ==== GrabExpress API (tracking pengiriman) ====
# Kredensial OAuth2 Client Credentials dari Grab Developer Portal.
GRAB_CLIENT_ID = os.getenv("GRAB_CLIENT_ID", "")
GRAB_CLIENT_SECRET = os.getenv("GRAB_CLIENT_SECRET", "")
# Base URL API. Sandbox: https://partner-api.grab.com/grab-express-sandbox
#                Production: https://partner-api.grab.com
GRAB_BASE_URL = os.getenv("GRAB_BASE_URL", "https://partner-api.grab.com")
# Endpoint OAuth token (biasanya https://api.grab.com/grabid/v1/oauth2/token)
GRAB_OAUTH_URL = os.getenv("GRAB_OAUTH_URL", "https://api.grab.com/grabid/v1/oauth2/token")
# Scope OAuth (opsional, tergantung konfigurasi partner Grab)
GRAB_SCOPE = os.getenv("GRAB_SCOPE", "")


def grab_configured() -> bool:
    return bool(GRAB_CLIENT_ID and GRAB_CLIENT_SECRET and GRAB_BASE_URL and GRAB_OAUTH_URL)


# ==== Eraspace Order (dumpdo) API ====
# URL endpoint dumpdo untuk cek status order & tracking.
ERASPACE_DUMPDO_URL = os.getenv("ERASPACE_DUMPDO_URL", "https://jeanne.eraspace.com/orders/v1/order/dumpdo")
# Nilai default parameter 'user' pada query string.
ERASPACE_USER = os.getenv("ERASPACE_USER", "erafone")
# Cookie Cloudflare (opsional). Isi kalau API menolak request tanpa cookie.
# Contoh: "__cf_bm=xxxxx". Cookie ini bisa expired, perlu diperbarui berkala.
ERASPACE_COOKIE = os.getenv("ERASPACE_COOKIE", "")

# ==== Eraspace Shipping Tracking PSD (awbpsd) API ====
AWBPSD_URL = os.getenv("AWBPSD_URL", "https://jeanne.eraspace.com/shippings/v2/tracking/order/oms/psd")
# Header Authorization (Basic). Default dari curl; bisa dioverride via env.
AWBPSD_AUTH = os.getenv(
    "AWBPSD_AUTH",
    "Basic c2hpcHBpbmdiYXNpYzo3NmNkNDJlZTQzZTUxNTIzZTAzNTVjZDE3NTMxY2ZjZjQxYjE2MWNmZDJjNTgwNDJkZjkxZTVmODU1MDQwYTQx",
)
AWBPSD_X_SOURCE = os.getenv("AWBPSD_X_SOURCE", "eraspace")
AWBPSD_X_PLATFORM = os.getenv("AWBPSD_X_PLATFORM", "omsservice")
# Cookie opsional (Cloudflare). Bisa expired, perlu diperbarui berkala.
AWBPSD_COOKIE = os.getenv("AWBPSD_COOKIE", "")

# ==== Stock Webhook Sync (External API) ====
# Endpoint untuk sync data stok ke sistem eksternal
WEBHOOK_STOCK_URL = os.getenv("WEBHOOK_STOCK_URL", "https://stockadapters.eraspace.com/v1/webhooks/stock")
# Cookie Cloudflare (opsional, bisa di-rotate via env)
WEBHOOK_STOCK_COOKIE = os.getenv("WEBHOOK_STOCK_COOKIE", "")

def webhook_stock_configured() -> bool:
    """Check jika webhook stock sudah dikonfigurasi."""
    return bool(WEBHOOK_STOCK_URL)

# ==== SQLLoader & Dynamic Scheduler ====
# Path folder untuk menyimpan file .sql yang akan dijalankan secara otomatis
SQL_FOLDER_PATH = os.getenv("SQL_FOLDER_PATH", "D:\\mybot\\tools\\queries")

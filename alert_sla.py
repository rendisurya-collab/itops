from playwright.sync_api import sync_playwright
import re
import requests
import time
from datetime import datetime
from collections import Counter

# ======================
# CONFIG
# ======================
TOKEN = "8471552922:AAG2PLHPrLorbz2Zk_Pkw6WlhVYrE4"
CHAT_ID = 553648540
URL = "https://servicedesk.erajaya.com/WOListView.do"

INTERVAL = 600  # 5 menit


# ======================
# TELEGRAM FUNCTION
# ======================
def send_telegram(msg):
    try:
        url = f"https://api.telegram.org/bot8471552922:AAG2PLHPrLorbz2Zk_Pkw6WlhVYrE4/sendMessage"
        requests.post(url, data={"chat_id": CHAT_ID, "text": msg}, timeout=10)
    except Exception as e:
        print("Telegram error:", e)


# ======================
# DASHBOARD CONSOLE
# ======================
def dashboard(
    open_status,
    investigation_status,
    transfer_status,
    waiting_status,
    pending_status,
    overdue_status,
    onhold_status,
    closed_status,
    assignee_info
):

    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    total_active = (
        open_status +
        investigation_status +
        transfer_status +
        waiting_status +
        pending_status +
        overdue_status +
        onhold_status
    )

    print("\n" + "=" * 50)
    print("TICKET MONITORING")
    print("=" * 50)
    print("Open                      :", open_status)
    print("In Progress Investigation :", investigation_status)
    print("Transfer L1               :", transfer_status)
    print("Waiting User Confirmation :", waiting_status)
    print("Pending                   :", pending_status)
    print("OverDue                   :", overdue_status)
    print("Onhold                    :", onhold_status)
    print("Closed                    :", closed_status)
    print()
    print("Total Active Ticket       :", total_active)

    if assignee_info:
        print("\nAssigned To:")
        for item in assignee_info:
            print(item)

    print()
    print("Last Check                :", now)
    print("=" * 50)


# ======================
# MAIN LOOP
# ======================
with sync_playwright() as p:

    browser = p.chromium.launch(headless=False)
    context = browser.new_context(storage_state="state.json")
    page = context.new_page()

    send_telegram("🟢 Ticket Monitoring STARTED")

    while True:
        try:
            # buka halaman
            page.goto(URL, timeout=15000)
            page.wait_for_timeout(5000)

            html = page.content()

            # ======================
            # STATUS COUNTER
            # ======================
            open_status = len(re.findall(r'title="Open"', html))
            investigation_status = len(re.findall(r'title="In Progress Investigation"', html))
            transfer_status = len(re.findall(r'title="Transfer L1"', html))
            waiting_status = len(re.findall(r'title="Waiting User Confirmation"', html))
            pending_status = len(re.findall(r'title="Pending"', html))
            overdue_status = len(re.findall(r'title="OverDue"', html))
            onhold_status = len(re.findall(r'title="Onhold"', html))
            closed_status = len(re.findall(r'title="Closed"', html))

            # ======================
            # ASSIGNEE
            # ======================

            team = [
                "Adelia Pebriani",
                "Bagus Dwi Susworo",
                "Bambang Purnomo Sidi",
                "Prizky Stefajar Darmaliz Gagah Utomo",
                "Rendi Surya Hadinata",
                "Teguh Wiguna",
                "Tri Sutrisno",
                "Unassigned"
            ]

            # Ambil technician dari kolom Assigned To
            technicians = re.findall(
                r'data-lv-action="technician".*?title="([^"]+)"',
                html,
                re.DOTALL
            )

            # DEBUG (hapus nanti jika sudah benar)
            print("\n=== TECHNICIANS ===")
            print(technicians[:20])

            counter = Counter(technicians)

            assignee_info = []

            for member in team:

                count = counter.get(member, 0)

                if count > 0:
                    assignee_info.append(
                        f"- {member} : {count}"
                    )

            # ======================
            # TOTAL ACTIVE
            # ======================
            total_active = (
                open_status +
                investigation_status +
                transfer_status +
                waiting_status +
                pending_status +
                overdue_status +
                onhold_status
            )

            # ======================
            # CONSOLE
            # ======================
            dashboard(
                open_status,
                investigation_status,
                transfer_status,
                waiting_status,
                pending_status,
                overdue_status,
                onhold_status,
                closed_status,
                assignee_info
            )

            # ======================
            # TELEGRAM MESSAGE
            # ======================
            msg = (
                "🚨 Ticket Update\n"
                "====================\n\n"
                f"Open                      : {open_status}\n"
                f"In Progress Investigation : {investigation_status}\n"
                f"Transfer L1               : {transfer_status}\n"
                f"Waiting User Confirmation : {waiting_status}\n"
                f"Pending                   : {pending_status}\n"
                f"OverDue                   : {overdue_status}\n"
                f"Onhold                    : {onhold_status}\n"
                f"Closed                    : {closed_status}\n\n"
                f"Total Active Ticket       : {total_active}\n"
            )

            if assignee_info:
                msg += "\nAssigned To:\n"
                msg += "\n".join(assignee_info)

            msg += (
                f"\n\nTime        : "
                f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
                "===================="
            )

            send_telegram(msg)

        except Exception as e:

            send_telegram(
                f"❌ ERROR MONITORING\n{str(e)}"
            )

        time.sleep(INTERVAL)
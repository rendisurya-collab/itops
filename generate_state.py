"""
Script untuk generate state.json (session login SDP).
Jalankan sekali saja: python generate_state.py

Browser akan terbuka → login manual ke SDP → setelah berhasil login,
tekan Enter di terminal → file state.json akan tersimpan.
"""
from playwright.sync_api import sync_playwright

URL = "https://servicedesk.erajaya.com"

with sync_playwright() as p:
    browser = p.chromium.launch(headless=False)
    context = browser.new_context()
    page = context.new_page()
    page.goto(URL)

    input("\n✅ Login manual di browser yang terbuka, lalu tekan ENTER di sini setelah berhasil login...\n")

    context.storage_state(path="state.json")
    browser.close()
    print("✅ state.json berhasil disimpan! Bot sekarang bisa menjalankan SLA Monitor.")

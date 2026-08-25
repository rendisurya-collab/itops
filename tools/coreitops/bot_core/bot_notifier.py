import os
import json
import time
import requests
import openpyxl
from bs4 import BeautifulSoup
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
from playwright.sync_api import sync_playwright

def load_config(config_path="config.json"):
    base_dir = os.path.dirname(os.path.abspath(__file__))
    full_path = os.path.join(base_dir, config_path)
    with open(full_path, "r", encoding="utf-8") as f:
        return json.load(f)

def load_sql_query(file_path):
    base_dir = os.path.dirname(os.path.abspath(__file__))
    full_path = os.path.join(base_dir, file_path)
    with open(full_path, "r", encoding="utf-8") as f:
        return f.read()

def send_telegram_document(bot_token, chat_id, file_path, caption=""):
    """Mengirimkan file Excel (.xlsx) sebagai dokumen ke Telegram."""
    url = f"https://api.telegram.org/bot{bot_token}/sendDocument"
    
    session = requests.Session()
    retries = Retry(total=3, backoff_factor=2, status_forcelist=[500, 502, 503, 504])
    session.mount("https://", HTTPAdapter(max_retries=retries))

    try:
        with open(file_path, "rb") as file_data:
            files = {"document": file_data}
            payload = {
                "chat_id": chat_id,
                "caption": caption
            }
            response = session.post(url, data=payload, files=files, timeout=60)
            response.raise_for_status()
            print(f"[INFO] File Excel '{os.path.basename(file_path)}' berhasil dikirim ke Telegram.")
    except Exception as e:
        print(f"[ERROR] Gagal mengirim file ke Telegram: {e}")

def create_excel_report(headers, data_list, output_path):
    """Membuat file Excel dari data query SQL."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Hasil Query"

    # Tulis Header
    if headers:
        ws.append(headers)

    # Tulis Data
    for row in data_list:
        ws.append(row)

    wb.save(output_path)

def run_adminer_playwright():
    config = load_config()
    adm = config["adminer"]
    tg = config["telegram"]
    
    # Membaca daftar file query & meratakannya jika berupa nested list
    raw_files = config.get("query_files", config.get("query_file", ["query.sql"]))
    if isinstance(raw_files, str):
        query_files = [raw_files]
    elif isinstance(raw_files, list):
        query_files = []
        for item in raw_files:
            if isinstance(item, list):
                query_files.extend(item)
            else:
                query_files.append(item)
    else:
        query_files = ["query.sql"]

    print(f"[INFO] Menemukan {len(query_files)} file query untuk dieksekusi: {query_files}")
    print("[INFO] Membuka browser Playwright...")
    
    with sync_playwright() as p:
        http_credentials = None
        if adm.get("basic_auth_user") and adm.get("basic_auth_pass"):
            http_credentials = {
                "username": adm["basic_auth_user"],
                "password": adm["basic_auth_pass"]
            }

        browser = p.chromium.launch(headless=True)
        context = browser.new_context(
            http_credentials=http_credentials,
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
        )
        page = context.new_page()

        try:
            print("[INFO] Menghubungkan ke Adminer...")
            page.goto(adm["url"], wait_until="networkidle")

            driver_select = page.locator("select[name='auth[driver]']")
            if driver_select.count() > 0 and adm.get("driver"):
                try:
                    driver_select.select_option(value=adm["driver"], timeout=2000)
                except Exception:
                    pass

            print("[INFO] Mengisi form login Adminer...")
            page.fill("input[name='auth[server]']", adm["server"])
            page.fill("input[name='auth[username]']", adm["username"])
            page.fill("input[name='auth[password]']", adm["password"])
            
            db_input = page.locator("input[name='auth[db]']")
            if db_input.count() > 0:
                db_input.fill(adm["db"])

            page.click("input[type='submit']")
            page.wait_for_load_state("networkidle")

            if page.locator("div.error").count() > 0:
                err_txt = page.locator("div.error").first.text_content()
                print(f"[ERROR] Gagal Login Adminer: {err_txt.strip()}")
                browser.close()
                return

            for q_file in query_files:
                print(f"\n==========================================")
                print(f"[INFO] Memproses file query: {q_file}")
                
                try:
                    sql_query = load_sql_query(q_file)
                except Exception as file_err:
                    print(f"[ERROR] Gagal membaca file {q_file}: {file_err}")
                    continue

                sql_link = page.locator("a:has-text('SQL command'), a:has-text('SQL dotaz'), a[href*='sql=']").first
                if sql_link.count() > 0:
                    sql_link.click()
                    page.wait_for_load_state("networkidle")
                else:
                    query_url = f"{adm['url']}?username={adm['username']}&db={adm['db']}&sql="
                    page.goto(query_url, wait_until="networkidle")

                page.wait_for_selector("textarea[name='query']", state="attached", timeout=10000)
                
                with page.expect_navigation(wait_until="networkidle"):
                    page.evaluate("""
                        (queryText) => {
                            const textarea = document.querySelector("textarea[name='query']");
                            if (textarea) {
                                textarea.value = queryText;
                                const form = textarea.closest("form");
                                if (form) {
                                    form.submit();
                                }
                            }
                        }
                    """, sql_query)

                html_content = page.content()
                soup = BeautifulSoup(html_content, "html.parser")

                # Cek jika ada error dari Adminer
                error_div = soup.find("div", class_="error")
                if error_div:
                    err_msg = error_div.text.strip()
                    print(f"[ERROR] Adminer Query Error ({q_file}): {err_msg}")
                    continue

                # Parsing Tabel
                table = soup.find("table", class_="printable") or soup.find("table")
                rows = table.find_all("tr") if table else []

                data_list = []
                headers = []
                if rows and len(rows) >= 2:
                    headers = [th.text.strip() for th in rows[0].find_all(["th", "td"])]
                    for row in rows[1:]:
                        cols = [td.text.strip() for td in row.find_all("td")]
                        if cols and any(cols):
                            data_list.append(cols)

                # Nama file excel temp
                base_name = os.path.splitext(q_file)[0]
                excel_filename = f"Hasil_{base_name}.xlsx"
                excel_filepath = os.path.join(os.path.dirname(__file__), excel_filename)

                # Jika tidak ada data (0 rows), buat Excel kosong atau dengan pesan
                if not data_list:
                    print(f"[INFO] Hasil query ({q_file}) kosong (0 baris). Mengirim Excel 'No Rows'...")
                    create_excel_report(["Status"], [["No Rows"]], excel_filepath)
                    caption = f"📊 HASIL QUERY ({q_file})\nStatus: No Rows"
                else:
                    # Buat file Excel dengan data lengkap
                    create_excel_report(headers, data_list, excel_filepath)
                    caption = f"📊 HASIL QUERY ({q_file})\nTotal Baris: {len(data_list)}"

                # Kirim file Excel ke Telegram
                send_telegram_document(tg["bot_token"], tg["chat_id"], excel_filepath, caption)

                # Hapus file Excel lokal setelah terkirim (Opsional)
                if os.path.exists(excel_filepath):
                    os.remove(excel_filepath)

                time.sleep(2)

            browser.close()
            print("\n[INFO] Semua file query selesai diproses.")

        except Exception as e:
            print(f"[ERROR] Terjadi kesalahan pada Playwright: {e}")
            browser.close()

if __name__ == "__main__":
    run_adminer_playwright()
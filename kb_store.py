"""Knowledge Base / Bank Data store — backed by Google Sheets.

Menyimpan FAQ aktif dan pertanyaan user yang belum terjawab (pending) di
Google Sheets (spreadsheet yang sama dengan ticket log), pada dua tab:

- Tab "FAQ"        : ID | Keywords | Question | Answer
- Tab "PendingFAQ" : ID | User ID | User Name | Question | Status | Answer | Created At | Answered At

Data tidak hilang saat Railway redeploy karena tersimpan di Google Sheets.

Fungsi publik tetap sama seperti versi JSON supaya bot.py tidak perlu diubah:
  add_faq, delete_faq, list_faqs, find_answer,
  add_pending, list_pending, get_pending, answer_pending, delete_pending
"""

import json
import logging
import datetime as dt

import config

logger = logging.getLogger(__name__)

FAQ_TAB = "FAQ"
PENDING_TAB = "PendingFAQ"
QUOTES_TAB = "Quotes"

FAQ_HEADER = ["ID", "Keywords", "Question", "Answer"]
PENDING_HEADER = ["ID", "User ID", "User Name", "Question", "Status", "Answer", "Created At", "Answered At"]
QUOTES_HEADER = ["Text", "Author"]

# Cache spreadsheet object supaya tidak auth berulang tiap operasi
_SPREADSHEET = None


def _now_str() -> str:
    return dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def _get_spreadsheet():
    """Return spreadsheet gspread (cached), atau None kalau credentials kosong."""
    global _SPREADSHEET
    if _SPREADSHEET is not None:
        return _SPREADSHEET

    creds_json = config.GOOGLE_SHEETS_CREDENTIALS
    sheet_id = config.GOOGLE_SHEETS_SPREADSHEET_ID
    if not creds_json or not sheet_id:
        logger.warning("Knowledge Base: GOOGLE_SHEETS credentials/spreadsheet ID kosong.")
        return None

    import gspread
    from google.oauth2.service_account import Credentials

    creds_dict = json.loads(creds_json)
    credentials = Credentials.from_service_account_info(
        creds_dict, scopes=["https://www.googleapis.com/auth/spreadsheets"]
    )
    gc = gspread.authorize(credentials)
    _SPREADSHEET = gc.open_by_key(sheet_id)
    return _SPREADSHEET


def _get_worksheet(tab: str, header: list):
    """Return worksheet, buat dengan header kalau belum ada."""
    ss = _get_spreadsheet()
    if ss is None:
        raise RuntimeError("Google Sheets credentials atau spreadsheet ID kosong.")
    try:
        ws = ss.worksheet(tab)
    except Exception:
        ws = ss.add_worksheet(title=tab, rows=1000, cols=len(header))
        ws.append_row(header, value_input_option="USER_ENTERED")
    return ws


def _faq_ws():
    return _get_worksheet(FAQ_TAB, FAQ_HEADER)


def _pending_ws():
    return _get_worksheet(PENDING_TAB, PENDING_HEADER)


def _next_id(rows: list, col_idx: int = 0) -> int:
    """rows = list baris (tanpa header). col_idx = index kolom ID."""
    max_id = 0
    for r in rows:
        if len(r) > col_idx and r[col_idx]:
            try:
                max_id = max(max_id, int(r[col_idx]))
            except (ValueError, TypeError):
                continue
    return max_id + 1


def _auto_keywords(text: str) -> list:
    """Ambil kata >3 huruf sebagai keyword default (buang kata sambung umum)."""
    stopwords = {
        "yang", "untuk", "dengan", "adalah", "dari", "pada", "atau", "dan",
        "apa", "apakah", "bagaimana", "kenapa", "mengapa", "dimana", "kapan",
        "gimana", "tolong", "mohon", "saya", "kami", "ini", "itu", "ada",
    }
    words = []
    for w in text.lower().replace("?", " ").replace(",", " ").replace(".", " ").split():
        w = w.strip()
        if len(w) > 3 and w not in stopwords:
            words.append(w)
    return words[:8]


def _tokenize(text: str) -> set:
    cleaned = text.lower().replace("?", " ").replace(",", " ").replace(".", " ").replace("!", " ")
    return {w for w in cleaned.split() if w}


# ---------------------------------------------------------------------------
# FAQ AKTIF
# ---------------------------------------------------------------------------

def list_faqs() -> list:
    """Return list of dict FAQ aktif."""
    try:
        ws = _faq_ws()
        rows = ws.get_all_values()
    except Exception as e:
        logger.error(f"Gagal baca FAQ dari Sheets: {e}")
        return []

    faqs = []
    for r in rows[1:]:  # skip header
        if len(r) < 4 or not r[0]:
            continue
        try:
            fid = int(r[0])
        except (ValueError, TypeError):
            continue
        keywords = [k.strip().lower() for k in r[1].split(",") if k.strip()]
        faqs.append({
            "id": fid,
            "keywords": keywords,
            "question": r[2].strip(),
            "answer": r[3].strip(),
        })
    return faqs


def add_faq(question: str, answer: str, keywords: list = None) -> dict:
    """Tambah FAQ baru ke tab FAQ."""
    ws = _faq_ws()
    rows = ws.get_all_values()
    new_id = _next_id(rows[1:], 0)

    if not keywords:
        keywords = _auto_keywords(question)
    kw_clean = [k.strip().lower() for k in keywords if k.strip()]

    ws.append_row(
        [new_id, ", ".join(kw_clean), question.strip(), answer.strip()],
        value_input_option="USER_ENTERED",
    )
    return {
        "id": new_id,
        "keywords": kw_clean,
        "question": question.strip(),
        "answer": answer.strip(),
    }


def delete_faq(faq_id: int) -> bool:
    ws = _faq_ws()
    rows = ws.get_all_values()
    for idx, r in enumerate(rows[1:], start=2):  # 1-based, skip header
        if r and r[0]:
            try:
                if int(r[0]) == int(faq_id):
                    ws.delete_rows(idx)
                    return True
            except (ValueError, TypeError):
                continue
    return False


# ---------------------------------------------------------------------------
# PENCARIAN (matching FAQ vs pertanyaan user)
# ---------------------------------------------------------------------------

_STOPWORDS = {
    "yang", "untuk", "dengan", "adalah", "dari", "pada", "atau", "dan",
    "apa", "apakah", "bagaimana", "gimana", "kenapa", "mengapa", "dimana",
    "kapan", "cara", "tolong", "mohon", "saya", "kami", "ini", "itu", "ada",
    "ke", "di", "ya", "dong", "sih", "nya", "aku", "kamu", "bisa", "boleh",
    "mau", "ingin", "biar", "supaya", "agar", "kalau", "jika", "bila",
    "gmn", "gimna", "solusi", "masalah", "problem",
}


def _meaningful_tokens(text: str) -> set:
    """Token setelah buang stopword — kata yang benar-benar bermakna."""
    return {t for t in _tokenize(text) if t not in _STOPWORDS and len(t) > 2}


def find_answer(query: str) -> dict:
    """Cari FAQ terbaik untuk pertanyaan user. Return dict FAQ atau None.

    Pencocokan berbasis KATA BERMAKNA (stopword seperti 'bagaimana cara'
    diabaikan) supaya pertanyaan generik tidak asal match.
    """
    query_lower = query.lower().strip()
    if len(query_lower) < 2:
        return None

    query_meaning = _meaningful_tokens(query_lower)
    # Kalau user tidak menuliskan satu pun kata bermakna (cuma stopword),
    # jangan match apa-apa.
    if not query_meaning:
        return None

    best = None
    best_score = 0

    for faq in list_faqs():
        score = 0
        q_faq = faq.get("question", "").lower().strip()
        faq_meaning = _meaningful_tokens(q_faq)

        # 1. Exact match penuh (persis sama) → poin tertinggi
        if q_faq and q_faq == query_lower:
            score += 10

        # 2. Overlap kata bermakna antara Question dan pertanyaan user.
        #    Dihitung sebagai rasio: berapa banyak kata bermakna FAQ yang
        #    tertutup oleh pertanyaan user. Butuh minimal separuh cocok.
        if faq_meaning:
            common = faq_meaning & query_meaning
            if common:
                coverage = len(common) / len(faq_meaning)
                if coverage >= 0.5:
                    score += 4 + len(common)  # makin banyak kata cocok, makin tinggi

        # 3. Match keyword (kata kunci). Keyword multi-kata dihitung kalau
        #    SEMUA kata bermakna keyword ada di pertanyaan user.
        for kw in faq.get("keywords", []):
            kw = kw.strip().lower()
            if not kw:
                continue
            kw_meaning = _meaningful_tokens(kw)
            if not kw_meaning:
                # keyword cuma stopword — abaikan supaya tidak asal match
                continue
            if kw in query_lower:
                score += 5
            elif kw_meaning.issubset(query_meaning):
                score += 4

        if score > best_score:
            best_score = score
            best = faq

    # Ambang minimal supaya hanya kecocokan yang cukup kuat yang dijawab
    return best if best_score >= 4 else None


# ---------------------------------------------------------------------------
# PERTANYAAN PENDING (belum terjawab)
# ---------------------------------------------------------------------------

def _read_pending_rows() -> list:
    """Return (worksheet, rows_all_values)."""
    ws = _pending_ws()
    return ws, ws.get_all_values()


def list_pending(only_unanswered: bool = True) -> list:
    try:
        _, rows = _read_pending_rows()
    except Exception as e:
        logger.error(f"Gagal baca PendingFAQ dari Sheets: {e}")
        return []

    result = []
    for r in rows[1:]:
        if len(r) < 5 or not r[0]:
            continue
        try:
            pid = int(r[0])
        except (ValueError, TypeError):
            continue
        status = r[4].strip() if len(r) > 4 else "unanswered"
        if only_unanswered and status != "unanswered":
            continue
        result.append({
            "id": pid,
            "user_id": r[1].strip() if len(r) > 1 else "",
            "user_name": r[2].strip() if len(r) > 2 else "",
            "question": r[3].strip() if len(r) > 3 else "",
            "status": status,
            "answer": r[5].strip() if len(r) > 5 else "",
            "created_at": r[6].strip() if len(r) > 6 else "",
        })
    return result


def get_pending(pending_id: int) -> dict:
    for p in list_pending(only_unanswered=False):
        if int(p.get("id", 0)) == int(pending_id):
            return p
    return None


def add_pending(question: str, user_id: str, user_name: str = "") -> dict:
    """Simpan pertanyaan user yang belum terjawab (hindari duplikat unanswered
    dari user & teks yang sama)."""
    ws, rows = _read_pending_rows()
    q_norm = question.strip().lower()

    for r in rows[1:]:
        if len(r) < 5:
            continue
        status = r[4].strip() if len(r) > 4 else ""
        r_uid = r[1].strip() if len(r) > 1 else ""
        r_q = r[3].strip().lower() if len(r) > 3 else ""
        if status == "unanswered" and r_uid == str(user_id) and r_q == q_norm:
            # sudah ada, jangan duplikat
            return {
                "id": int(r[0]) if r[0] else 0,
                "user_id": r_uid,
                "user_name": r[2].strip() if len(r) > 2 else "",
                "question": question.strip(),
                "status": "unanswered",
            }

    new_id = _next_id(rows[1:], 0)
    created = _now_str()
    ws.append_row(
        [new_id, str(user_id), user_name, question.strip(), "unanswered", "", created, ""],
        value_input_option="USER_ENTERED",
    )
    return {
        "id": new_id,
        "user_id": str(user_id),
        "user_name": user_name,
        "question": question.strip(),
        "status": "unanswered",
        "created_at": created,
    }


def answer_pending(pending_id: int, answer: str) -> dict:
    """Tandai pending answered + promosikan jadi FAQ aktif.

    Return {"pending": ..., "faq": ...} atau None kalau id tidak ada.
    """
    ws, rows = _read_pending_rows()
    target_row_idx = None
    target = None

    for idx, r in enumerate(rows[1:], start=2):  # 1-based, skip header
        if r and r[0]:
            try:
                if int(r[0]) == int(pending_id):
                    target_row_idx = idx
                    target = {
                        "id": int(r[0]),
                        "user_id": r[1].strip() if len(r) > 1 else "",
                        "user_name": r[2].strip() if len(r) > 2 else "",
                        "question": r[3].strip() if len(r) > 3 else "",
                        "status": r[4].strip() if len(r) > 4 else "",
                    }
                    break
            except (ValueError, TypeError):
                continue

    if target is None:
        return None

    answered_at = _now_str()
    # Update kolom Status (E), Answer (F), Answered At (H)
    ws.update(f"E{target_row_idx}", [["answered"]])
    ws.update(f"F{target_row_idx}", [[answer.strip()]])
    ws.update(f"H{target_row_idx}", [[answered_at]])

    target["status"] = "answered"
    target["answer"] = answer.strip()
    target["answered_at"] = answered_at

    # Promosikan ke FAQ aktif
    faq_entry = add_faq(target["question"], answer, _auto_keywords(target["question"]))
    return {"pending": target, "faq": faq_entry}


def delete_pending(pending_id: int) -> bool:
    ws, rows = _read_pending_rows()
    for idx, r in enumerate(rows[1:], start=2):
        if r and r[0]:
            try:
                if int(r[0]) == int(pending_id):
                    ws.delete_rows(idx)
                    return True
            except (ValueError, TypeError):
                continue
    return False


# ---------------------------------------------------------------------------
# QUOTES (tab Quotes: Text | Author)
# ---------------------------------------------------------------------------

def _quotes_ws():
    return _get_worksheet(QUOTES_TAB, QUOTES_HEADER)


def list_quotes() -> list:
    """Return list of dict {"text", "author"} dari tab Quotes."""
    ws = _quotes_ws()
    rows = ws.get_all_values()
    quotes = []
    for r in rows[1:]:  # skip header
        if not r or not r[0].strip():
            continue
        quotes.append({
            "text": r[0].strip(),
            "author": (r[1].strip() if len(r) > 1 and r[1].strip() else "Anonim"),
        })
    return quotes


def add_quote(text: str, author: str = "Anonim") -> dict:
    """Tambah quote baru ke tab Quotes."""
    ws = _quotes_ws()
    ws.append_row([text.strip(), (author or "Anonim").strip()], value_input_option="USER_ENTERED")
    return {"text": text.strip(), "author": (author or "Anonim").strip()}


def random_quote() -> dict:
    """Ambil 1 quote acak dari tab Quotes, atau None kalau kosong/gagal."""
    import random
    try:
        quotes = list_quotes()
    except Exception as e:
        logger.error(f"Gagal baca Quotes dari Sheets: {e}")
        return None
    if not quotes:
        return None
    return random.choice(quotes)

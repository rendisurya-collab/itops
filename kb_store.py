"""Knowledge Base / Bank Data store.

Menyimpan FAQ aktif dan pertanyaan user yang belum terjawab (pending) di
file JSON `knowledge_base.json`. Selalu baca ulang dari file supaya perubahan
manual lewat text editor langsung kepakai tanpa restart bot.

Struktur file:
{
  "active_faqs": [
    {"id": 1, "keywords": [...], "question": "...", "answer": "..."}
  ],
  "pending_questions": [
    {"id": 101, "user_id": "...", "user_name": "...", "question": "...",
     "status": "unanswered", "created_at": "..."}
  ]
}
"""

import json
import os
import datetime as dt

KB_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "knowledge_base.json")

_DEFAULT = {"active_faqs": [], "pending_questions": []}


def _ensure_file():
    if not os.path.exists(KB_FILE):
        with open(KB_FILE, "w", encoding="utf-8") as f:
            json.dump(_DEFAULT, f, ensure_ascii=False, indent=2)


def load_kb() -> dict:
    """Baca seluruh knowledge base dari file."""
    _ensure_file()
    try:
        with open(KB_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
    except (json.JSONDecodeError, FileNotFoundError):
        return {"active_faqs": [], "pending_questions": []}

    if not isinstance(data, dict):
        return {"active_faqs": [], "pending_questions": []}
    data.setdefault("active_faqs", [])
    data.setdefault("pending_questions", [])
    return data


def save_kb(data: dict):
    with open(KB_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def _next_id(items: list) -> int:
    max_id = 0
    for item in items:
        try:
            max_id = max(max_id, int(item.get("id", 0)))
        except (ValueError, TypeError):
            continue
    return max_id + 1


def _now_str() -> str:
    return dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")


# ---------------------------------------------------------------------------
# FAQ AKTIF
# ---------------------------------------------------------------------------

def add_faq(question: str, answer: str, keywords: list = None) -> dict:
    """Tambah FAQ baru ke active_faqs. Keywords auto dari pertanyaan bila kosong."""
    data = load_kb()
    if not keywords:
        # Ambil kata bermakna dari pertanyaan sebagai keyword default
        keywords = _auto_keywords(question)
    entry = {
        "id": _next_id(data["active_faqs"]),
        "keywords": [k.strip().lower() for k in keywords if k.strip()],
        "question": question.strip(),
        "answer": answer.strip(),
    }
    data["active_faqs"].append(entry)
    save_kb(data)
    return entry


def delete_faq(faq_id: int) -> bool:
    data = load_kb()
    before = len(data["active_faqs"])
    data["active_faqs"] = [f for f in data["active_faqs"] if int(f.get("id", 0)) != int(faq_id)]
    if len(data["active_faqs"]) == before:
        return False
    save_kb(data)
    return True


def list_faqs() -> list:
    return load_kb()["active_faqs"]


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


# ---------------------------------------------------------------------------
# PENCARIAN (matching FAQ vs pertanyaan user)
# ---------------------------------------------------------------------------

def _tokenize(text: str) -> set:
    cleaned = text.lower().replace("?", " ").replace(",", " ").replace(".", " ").replace("!", " ")
    return {w for w in cleaned.split() if w}


def find_answer(query: str) -> dict:
    """Cari FAQ terbaik untuk pertanyaan user.

    Return dict FAQ yang cocok (dengan skor tertinggi), atau None kalau tidak ada
    yang cukup relevan. Kombinasi substring match + token overlap sederhana.
    """
    query_lower = query.lower().strip()
    if len(query_lower) < 2:
        return None

    query_tokens = _tokenize(query_lower)
    best = None
    best_score = 0

    for faq in load_kb()["active_faqs"]:
        score = 0

        # 1. Match pertanyaan (substring dua arah)
        q_faq = faq.get("question", "").lower().strip()
        if q_faq:
            if q_faq == query_lower:
                score += 10  # exact match
            elif q_faq in query_lower or query_lower in q_faq:
                score += 5

        # 2. Match keyword (frasa) — substring
        for kw in faq.get("keywords", []):
            kw = kw.strip().lower()
            if not kw:
                continue
            if kw in query_lower:
                score += 4
            else:
                # token overlap: semua kata di keyword ada di pertanyaan user
                kw_tokens = _tokenize(kw)
                if kw_tokens and kw_tokens.issubset(query_tokens):
                    score += 3

        # 3. Token overlap pertanyaan FAQ vs user
        faq_tokens = _tokenize(q_faq)
        if faq_tokens:
            overlap = len(faq_tokens & query_tokens)
            if overlap >= 2:
                score += overlap

        if score > best_score:
            best_score = score
            best = faq

    # Ambang minimal supaya tidak asal match
    return best if best_score >= 3 else None


# ---------------------------------------------------------------------------
# PERTANYAAN PENDING (belum terjawab)
# ---------------------------------------------------------------------------

def add_pending(question: str, user_id: str, user_name: str = "") -> dict:
    """Simpan pertanyaan user yang belum terjawab. Hindari duplikat teks yang sama
    dari user yang sama dan masih unanswered."""
    data = load_kb()
    q_norm = question.strip().lower()

    for p in data["pending_questions"]:
        if (
            p.get("status") == "unanswered"
            and str(p.get("user_id")) == str(user_id)
            and p.get("question", "").strip().lower() == q_norm
        ):
            return p  # sudah ada, jangan duplikat

    entry = {
        "id": _next_id(data["pending_questions"]),
        "user_id": str(user_id),
        "user_name": user_name,
        "question": question.strip(),
        "status": "unanswered",
        "created_at": _now_str(),
    }
    data["pending_questions"].append(entry)
    save_kb(data)
    return entry


def list_pending(only_unanswered: bool = True) -> list:
    items = load_kb()["pending_questions"]
    if only_unanswered:
        return [p for p in items if p.get("status") == "unanswered"]
    return items


def get_pending(pending_id: int) -> dict:
    for p in load_kb()["pending_questions"]:
        if int(p.get("id", 0)) == int(pending_id):
            return p
    return None


def answer_pending(pending_id: int, answer: str) -> dict:
    """Jawab pertanyaan pending: tandai answered DAN promosikan jadi FAQ aktif.

    Return dict berisi {"pending": ..., "faq": ...} atau None kalau id tidak ada.
    """
    data = load_kb()
    target = None
    for p in data["pending_questions"]:
        if int(p.get("id", 0)) == int(pending_id):
            target = p
            break
    if target is None:
        return None

    target["status"] = "answered"
    target["answer"] = answer.strip()
    target["answered_at"] = _now_str()

    # Promosikan ke FAQ aktif
    faq_entry = {
        "id": _next_id(data["active_faqs"]),
        "keywords": _auto_keywords(target["question"]),
        "question": target["question"].strip(),
        "answer": answer.strip(),
    }
    data["active_faqs"].append(faq_entry)
    save_kb(data)
    return {"pending": target, "faq": faq_entry}


def delete_pending(pending_id: int) -> bool:
    data = load_kb()
    before = len(data["pending_questions"])
    data["pending_questions"] = [
        p for p in data["pending_questions"] if int(p.get("id", 0)) != int(pending_id)
    ]
    if len(data["pending_questions"]) == before:
        return False
    save_kb(data)
    return True

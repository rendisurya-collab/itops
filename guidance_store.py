import json
import os

GUIDANCE_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "guidance.json")


def _ensure_file():
    if not os.path.exists(GUIDANCE_FILE):
        with open(GUIDANCE_FILE, "w", encoding="utf-8") as f:
            json.dump([], f, ensure_ascii=False, indent=2)


def load_guidance() -> list:
    """Selalu baca ulang dari file, supaya perubahan manual lewat text editor
    langsung kepakai tanpa perlu restart bot."""
    _ensure_file()
    try:
        with open(GUIDANCE_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
            return data if isinstance(data, list) else []
    except (json.JSONDecodeError, FileNotFoundError):
        return []


def save_guidance(data: list):
    with open(GUIDANCE_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def _next_id(data: list) -> str:
    max_id = 0
    for item in data:
        try:
            max_id = max(max_id, int(item.get("id", "0")))
        except (ValueError, TypeError):
            continue
    return str(max_id + 1).zfill(4)


def add_guidance(title: str, keywords: list, content: str, attachment: dict = None, action: dict = None) -> dict:
    data = load_guidance()
    entry = {
        "id": _next_id(data),
        "title": title.strip(),
        "keywords": [k.strip().lower() for k in keywords if k.strip()],
        "content": content.strip(),
        "attachment": attachment,
        "action": action,
    }
    data.append(entry)
    save_guidance(data)
    return entry


def get_guidance(guidance_id: str) -> dict:
    for item in load_guidance():
        if item.get("id") == guidance_id:
            return item
    return None


def update_guidance(guidance_id: str, updates: dict) -> dict:
    """Update field-field tertentu pada satu guidance. `updates` cuma berisi
    key yang memang mau diubah."""
    data = load_guidance()
    for item in data:
        if item.get("id") == guidance_id:
            item.update(updates)
            save_guidance(data)
            return item
    return None


def delete_guidance(guidance_id: str) -> bool:
    data = load_guidance()
    new_data = [d for d in data if d.get("id") != guidance_id]
    if len(new_data) == len(data):
        return False
    save_guidance(new_data)
    return True


def find_matches(query: str) -> list:
    """Cari guidance yang cocok berdasarkan judul & kata kunci.
    Match ke kata kunci diprioritaskan (skor lebih tinggi) daripada match ke judul."""
    query_lower = query.lower().strip()
    if len(query_lower) < 3:
        return []
    scored = []
    for item in load_guidance():
        score = 0
        title_lower = item.get("title", "").lower()
        if title_lower and (title_lower in query_lower or query_lower in title_lower):
            score += 2
        for kw in item.get("keywords", []):
            if kw and (kw in query_lower or query_lower in kw):
                score += 3
                break
        if score > 0:
            scored.append((score, item))
    scored.sort(key=lambda pair: pair[0], reverse=True)
    return [item for _, item in scored]

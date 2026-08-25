import json
import os

ACCOUNTS_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "jira_accounts.json")


def _ensure_file():
    if not os.path.exists(ACCOUNTS_FILE):
        with open(ACCOUNTS_FILE, "w", encoding="utf-8") as f:
            json.dump({}, f)


def load_accounts() -> dict:
    _ensure_file()
    try:
        with open(ACCOUNTS_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
            return data if isinstance(data, dict) else {}
    except (json.JSONDecodeError, FileNotFoundError):
        return {}


def save_accounts(data: dict):
    with open(ACCOUNTS_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def get_account(user_id) -> dict:
    return load_accounts().get(str(user_id))


def set_account(user_id, email: str, api_token: str, display_name: str = ""):
    data = load_accounts()
    data[str(user_id)] = {
        "email": email.strip(),
        "api_token": api_token.strip(),
        "display_name": display_name.strip(),
    }
    save_accounts(data)


def remove_account(user_id) -> bool:
    data = load_accounts()
    if str(user_id) in data:
        del data[str(user_id)]
        save_accounts(data)
        return True
    return False

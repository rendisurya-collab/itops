"""
Handler untuk /synctowebhook - Sinkronisasi data stock ke webhook eksternal API.
"""
import re
import json
import datetime as dt
import logging
from typing import Dict

try:
    import cloudscraper
    HAS_CLOUDSCRAPER = True
except ImportError:
    HAS_CLOUDSCRAPER = False

import requests

logger = logging.getLogger(__name__)


def parse_synctowebhook_text(text: str) -> Dict:
    """
    Parse input text untuk sync webhook.
    
    Support format:
    - Shortcut: "article_code": ABC123, "site_code": SS20, ...
    - Label: article_code: ABC123\nsite_code: SS20\n...
    """
    data = {}
    
    # Try parse as JSON-like (dengan quotes)
    if '"' in text:
        try:
            # Clean up: "field": value → field: value
            clean = re.sub(r'"([^"]+)"\s*:\s*', r'\1: ', text)
            # Parse comma-separated
            for item in clean.split(','):
                item = item.strip()
                if ':' in item:
                    key, val = item.split(':', 1)
                    data[key.strip().lower()] = val.strip()
        except Exception as e:
            logger.debug(f"Error parsing JSON-like format: {e}")
    
    # Parse sebagai label format (line by line)
    if not data:
        for line in text.split('\n'):
            line = line.strip()
            if not line or ':' not in line:
                continue
            try:
                key, val = line.split(':', 1)
                data[key.strip().lower()] = val.strip()
            except ValueError:
                continue
    
    return data


def sync_stock_to_webhook(
    article_code: str,
    site_code: str,
    company_code: str,
    stock: int,
    timestamps: str,
    webhook_url: str,
    webhook_cookie: str = None,
) -> Dict:
    """
    Kirim data stock ke webhook eksternal API.
    
    Args:
        article_code: Kode artikel/SKU
        site_code: Kode lokasi/site
        company_code: Kode perusahaan
        stock: Jumlah stok (integer)
        timestamps: ISO-8601 timestamp (e.g., 2026-09-02T23:05:00Z)
        webhook_url: URL endpoint webhook
        webhook_cookie: Optional Cloudflare cookie
    
    Return: 
        {
            "success": bool,
            "status_code": int,
            "error": str,
            "response": dict (jika success)
        }
    """
    
    # Build payload
    payload = {
        "article_code": article_code,
        "site_code": site_code,
        "company_code": company_code,
        "stock": stock,
        "timestamps": timestamps,
    }
    
    # Build headers
    headers = {
        "Content-Type": "application/json",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        "Accept": "application/json",
        "Accept-Language": "en-US,en;q=0.9",
        "Referer": "https://stockadapters.eraspace.com/",
    }
    
    if webhook_cookie:
        headers["Cookie"] = webhook_cookie
    
    try:
        # Try dengan cloudscraper jika available (untuk bypass Cloudflare otomatis)
        if HAS_CLOUDSCRAPER and "cloudflare" in webhook_url.lower():
            logger.info(f"Using cloudscraper to bypass Cloudflare")
            scraper = cloudscraper.create_scraper()
            response = scraper.post(
                webhook_url,
                headers=headers,
                json=payload,
                timeout=30,
            )
        else:
            # Fallback ke requests biasa
            response = requests.post(
                webhook_url,
                headers=headers,
                json=payload,
                timeout=30,
            )
        
        # Log response
        logger.info(
            f"Webhook sync response: status={response.status_code}, "
            f"article={article_code}, stock={stock}"
        )
        
        if response.status_code in [200, 201]:
            # Success
            try:
                resp_json = response.json()
            except:
                resp_json = {}
            
            return {
                "success": True,
                "status_code": response.status_code,
                "error": None,
                "response": resp_json,
            }
        else:
            # Error response from API
            # Truncate HTML responses (Cloudflare error pages)
            error_text = response.text[:500] if len(response.text) < 1000 else f"[{len(response.text)} chars HTML response]"
            return {
                "success": False,
                "status_code": response.status_code,
                "error": error_text,
                "response": None,
            }
    
    except requests.exceptions.Timeout:
        return {
            "success": False,
            "status_code": None,
            "error": "Request timeout (30 detik)",
            "response": None,
        }
    except requests.exceptions.ConnectionError as e:
        return {
            "success": False,
            "status_code": None,
            "error": f"Connection error: {str(e)[:200]}",
            "response": None,
        }
    except Exception as e:
        logger.exception(f"Webhook sync error: {e}")
        return {
            "success": False,
            "status_code": None,
            "error": str(e)[:200],
            "response": None,
        }


def auto_fill_timestamp(ts: str) -> str:
    """Auto-fill timestamp jika kosong atau placeholder.
    
    Placeholder: "(datenow)", "now", "", "datenow"
    Return: ISO-8601 UTC timestamp
    """
    if not ts or ts.lower() in ["(datenow)", "now", "datenow"]:
        return dt.datetime.now(dt.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    return ts

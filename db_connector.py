"""
Database Connector untuk QueryExecutor - support multiple databases via Adminer (Playwright-based).
Menggunakan logic yang sama seperti /query command di bot.py.
"""
import json
import logging
from pathlib import Path
from typing import Dict, List, Optional, Tuple
import os
import subprocess

logger = logging.getLogger(__name__)


def ensure_playwright_browsers():
    """
    Ensure Playwright browsers are installed.
    Uses playwright._impl._install module to install browsers programmatically.
    """
    import sys
    
    try:
        # Try importing - if it fails, browsers aren't installed
        from playwright.sync_api import sync_playwright
        logger.info("Playwright browsers check passed")
        return True
    except Exception as e:
        logger.warning(f"Playwright browsers check failed: {e}")
    
    # Attempt to install
    try:
        logger.info("Attempting to install Playwright browsers...")
        
        # Method 1: Use subprocess with python -m
        result = subprocess.run(
            [sys.executable, "-m", "playwright", "install", "chromium"],
            capture_output=True,
            text=True,
            timeout=600
        )
        
        if result.returncode == 0:
            logger.info(f"✓ Playwright chromium installed successfully")
            return True
        else:
            logger.error(f"Playwright install failed with code {result.returncode}")
            logger.error(f"STDERR: {result.stderr}")
            logger.error(f"STDOUT: {result.stdout}")
            
            # Retry with direct CLI
            logger.warning("Retrying with direct CLI...")
            result2 = subprocess.run(
                ["playwright", "install", "chromium"],
                capture_output=True,
                text=True,
                timeout=600
            )
            if result2.returncode == 0:
                logger.info(f"✓ Playwright chromium installed via CLI")
                return True
            else:
                logger.error(f"Playwright CLI failed: {result2.stderr}")
                return False
                
    except subprocess.TimeoutExpired:
        logger.error("Playwright installation timeout (600s)")
        return False
    except Exception as e:
        logger.error(f"Error installing Playwright browsers: {e}", exc_info=True)
        return False


class DatabaseConnector:
    """
    Database connector yang support multiple databases via Adminer (Playwright automation).
    Sama seperti logic /query command di bot.py.
    """

    def __init__(self, config_file: str):
        """
        Initialize DatabaseConnector.

        Args:
            config_file: Path ke config.json (misal: D:\\mybot\\tools\\coreitops\\bot_core\\config.json)
        """
        self.config_file = Path(config_file)
        self.config: Dict = {}

        if self.config_file.exists():
            self.load_config()

    def load_config(self) -> bool:
        """
        Load config dari file.

        Return:
            True jika berhasil
        """
        try:
            with open(self.config_file, 'r', encoding='utf-8') as f:
                self.config = json.load(f)
            logger.info(f"Config loaded: {len(self.config.get('databases', []))} databases")
            return True
        except Exception as e:
            logger.error(f"Error loading config: {e}")
            return False

    def execute_query(self, db_name: str, query: str) -> Tuple[bool, list, str]:
        """
        Execute query ke database tertentu via Adminer (Playwright).

        Args:
            db_name: Nama database
            query: SQL query string

        Return:
            (success, rows, error_message)
        """
        try:
            # Ensure browsers installed before trying to use them
            install_ok = ensure_playwright_browsers()
            if not install_ok:
                logger.warning("Playwright installation may have failed, attempting to continue anyway...")
            
            from playwright.sync_api import sync_playwright
            from bs4 import BeautifulSoup
            
            # Find database config
            db_config = None
            for db in self.config.get('databases', []):
                if db.get('name') == db_name:
                    db_config = db.get('adminer', {})
                    break

            if not db_config:
                return False, [], f"Database '{db_name}' tidak ditemukan di config"

            # Get Adminer credentials
            adm = db_config
            http_credentials = None
            if adm.get('basic_auth_user') and adm.get('basic_auth_pass'):
                http_credentials = {
                    'username': adm['basic_auth_user'],
                    'password': adm['basic_auth_pass']
                }

            rows = []

            with sync_playwright() as p:
                browser = p.chromium.launch(headless=True)
                ctx = browser.new_context(
                    http_credentials=http_credentials,
                    user_agent='Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
                )
                page = ctx.new_page()

                try:
                    # Login to Adminer
                    page.goto(adm['url'], wait_until='networkidle', timeout=30000)

                    driver_select = page.locator("select[name='auth[driver]']")
                    if driver_select.count() > 0 and adm.get('driver'):
                        try:
                            driver_select.select_option(value=adm['driver'], timeout=2000)
                        except Exception:
                            pass

                    page.fill("input[name='auth[server]']", adm['server'])
                    page.fill("input[name='auth[username]']", adm['username'])
                    page.fill("input[name='auth[password]']", adm['password'])

                    db_input = page.locator("input[name='auth[db]']")
                    if db_input.count() > 0:
                        db_input.fill(adm['db'])

                    page.click("input[type='submit']")
                    page.wait_for_load_state('networkidle', timeout=30000)

                    if page.locator('div.error').count() > 0:
                        err_txt = page.locator('div.error').first.text_content()
                        return False, [], f"Adminer login error: {err_txt.strip()}"

                    # Execute query
                    # Navigate to SQL query page
                    page.goto(adm['url'] + '?sql=', wait_until='networkidle', timeout=30000)

                    # Fill query
                    query_textarea = page.locator("textarea[name='query']")
                    if query_textarea.count() == 0:
                        return False, [], "Query textarea tidak ditemukan di Adminer"

                    query_textarea.fill(query)

                    # Execute
                    page.click("input[type='submit']")
                    page.wait_for_load_state('networkidle', timeout=60000)

                    # Parse results
                    # Look for result table
                    result_table = page.locator('table')
                    if result_table.count() == 0:
                        logger.warning(f"Tidak ada result table untuk query '{query[:50]}...'")
                        return True, [], None

                    # Get table HTML
                    table_html = result_table.first.inner_html()
                    soup = BeautifulSoup(table_html, 'html.parser')

                    # Parse rows
                    rows_html = soup.find_all('tr')
                    if len(rows_html) <= 1:  # Only header
                        return True, [], None

                    # Skip header row
                    for row_html in rows_html[1:]:
                        cells = row_html.find_all('td')
                        if cells:
                            row_data = tuple(cell.text.strip() for cell in cells)
                            rows.append(row_data)

                    logger.info(f"Query executed via Adminer: {len(rows)} rows returned from '{db_name}'")
                    return True, rows, None

                except Exception as e:
                    error_msg = f"Query execution error: {str(e)}"
                    logger.error(error_msg)
                    return False, [], error_msg
                finally:
                    try:
                        page.close()
                        ctx.close()
                        browser.close()
                    except:
                        pass

        except Exception as e:
            error_msg = f"Database connection error: {str(e)}"
            logger.error(error_msg)
            return False, [], error_msg

    def get_database_for_query(self, query_name: str) -> Optional[str]:
        """
        Find database name untuk query tertentu.

        Args:
            query_name: Nama query file (tanpa .sql)

        Return:
            Database name atau None jika tidak ditemukan
        """
        for db in self.config.get('databases', []):
            query_files = db.get('query_files', [])
            if isinstance(query_files, str):
                query_files = [query_files]
            
            # Flatten if nested list
            flat_files = []
            for item in query_files:
                if isinstance(item, list):
                    flat_files.extend(item)
                else:
                    flat_files.append(item)
            
            # Check apakah query_name match dengan salah satu query file
            for qf in flat_files:
                if qf.replace('.sql', '') == query_name or qf == f"{query_name}.sql":
                    return db.get('name')
        return None


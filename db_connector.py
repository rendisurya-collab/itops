"""
Database Connector untuk QueryExecutor - support multiple databases via Adminer (Playwright-based).
Menggunakan logic yang sama seperti /query command di bot.py.
"""
import json
import logging
from pathlib import Path
from typing import Dict, List, Optional, Tuple
import os

logger = logging.getLogger(__name__)


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
                    logger.info(f"Navigating to Adminer: {adm['url']}")
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

                    logger.info("Submitting login form...")
                    page.click("input[type='submit']")
                    page.wait_for_load_state('networkidle', timeout=30000)

                    if page.locator('div.error').count() > 0:
                        err_txt = page.locator('div.error').first.text_content()
                        return False, [], f"Adminer login error: {err_txt.strip()}"

                    logger.info(f"Login successful, navigating to SQL page...")
                    # Execute query - navigate to SQL query page
                    # Try multiple possible URLs for SQL page
                    sql_urls = [
                        adm['url'] + '?sql=',
                        adm['url'].rstrip('/') + '/?sql=',
                        adm['url'].rstrip('/') + '/?action=sql',
                    ]
                    
                    page_loaded = False
                    for sql_url in sql_urls:
                        try:
                            logger.info(f"Trying SQL URL: {sql_url}")
                            page.goto(sql_url, wait_until='load', timeout=15000)
                            
                            # Log page info for debugging
                            page_title = page.title()
                            page_url = page.url
                            logger.info(f"Page title: {page_title}, URL: {page_url}")
                            
                            # Check if we got redirected to login (common issue)
                            if 'login' in page_title.lower() or 'auth' in page_title.lower():
                                logger.warning(f"Seems we got redirected to login page. Title: {page_title}")
                                continue
                            
                            # Wait for textarea with longer timeout and allow DOM mutations
                            try:
                                query_textarea = page.locator("textarea[name='query']")
                                # Wait for it to be visible
                                query_textarea.first.wait_for(timeout=5000, state='visible')
                                logger.info(f"Found query textarea at {sql_url}")
                                page_loaded = True
                                break
                            except Exception as wait_error:
                                logger.warning(f"Textarea not visible at {sql_url}: {wait_error}")
                                
                                # Try generic textarea selector
                                all_textareas = page.locator("textarea")
                                logger.info(f"Found {all_textareas.count()} textareas (trying generic selector)")
                                if all_textareas.count() > 0:
                                    # Found textarea with generic selector
                                    logger.info("Using first textarea found")
                                    page_loaded = True
                                    break
                                    
                        except Exception as e:
                            logger.warning(f"SQL URL {sql_url} failed: {e}")
                            continue
                    
                    if not page_loaded:
                        # Log page content for debugging
                        page_content = page.content()
                        logger.error(f"Could not find SQL page. Current URL: {page.url}")
                        logger.error(f"Page title: {page.title()}")
                        logger.error(f"Page length: {len(page_content)}")
                        
                        # Save page content to file for analysis
                        import time
                        timestamp = int(time.time())
                        debug_file = f"/tmp/adminer_debug_{timestamp}.html"
                        try:
                            with open(debug_file, 'w', encoding='utf-8') as f:
                                f.write(page_content)
                            logger.info(f"Page content saved to {debug_file}")
                        except Exception as e:
                            logger.warning(f"Could not save debug file: {e}")
                        
                        # Try to find textarea with any name
                        all_textareas = page.locator("textarea")
                        logger.error(f"Found {all_textareas.count()} textareas on page")
                        for i in range(all_textareas.count()):
                            ta_name = all_textareas.nth(i).get_attribute("name")
                            logger.error(f"Textarea {i}: name={ta_name}")
                        
                        # Try to find any input or form elements
                        all_inputs = page.locator("input[type='text'], input[type='submit'], textarea")
                        logger.error(f"Found {all_inputs.count()} form elements total")
                        
                        return False, [], "Query textarea tidak ditemukan di Adminer - SQL page tidak accessible"

                    # Fill query
                    query_textarea = page.locator("textarea[name='query']")
                    if query_textarea.count() == 0:
                        # Try generic textarea selector
                        query_textarea = page.locator("textarea").first
                    
                    query_textarea.fill(query)
                    logger.info(f"Query filled, executing...")

                    # Execute - find submit button
                    submit_btn = page.locator("input[type='submit']").first
                    submit_btn.click()
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
                    logger.error(error_msg, exc_info=True)
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
            logger.error(error_msg, exc_info=True)
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


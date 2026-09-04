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

                    logger.info(f"Login successful, current URL: {page.url}")
                    
                    # After login, we're already in the database
                    # Now we need to execute query - but can't navigate (loses session)
                    # Instead, look for query form on current page or in sidebar
                    
                    # Try to find query textarea on current page
                    query_textarea = page.locator("textarea[name='query']")
                    
                    if query_textarea.count() == 0:
                        # Try to click on SQL/Query link/menu if available
                        logger.info("Looking for SQL menu option...")
                        
                        # Try common SQL menu selectors
                        sql_menu_selectors = [
                            "a:has-text('SQL')",
                            "a:has-text('SQL command')",
                            "a:has-text('Execute SQL')",
                            "button:has-text('SQL')",
                        ]
                        
                        menu_found = False
                        for selector in sql_menu_selectors:
                            try:
                                menu_item = page.locator(selector)
                                if menu_item.count() > 0:
                                    logger.info(f"Found SQL menu: {selector}")
                                    menu_item.first.click()
                                    page.wait_for_load_state('load', timeout=10000)
                                    menu_found = True
                                    break
                            except Exception as e:
                                logger.warning(f"Menu selector {selector} failed: {e}")
                        
                        # Check if we got redirected to login again after clicking SQL
                        page_title = page.title()
                        if 'login' in page_title.lower():
                            logger.info(f"Got redirected to login after SQL click. Re-login for DB access...")
                            # Need to login again with DB credentials
                            
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

                            logger.info("Submitting login form again...")
                            page.click("input[type='submit']")
                            page.wait_for_load_state('networkidle', timeout=30000)
                            
                            if page.locator('div.error').count() > 0:
                                err_txt = page.locator('div.error').first.text_content()
                                return False, [], f"Adminer re-login error: {err_txt.strip()}"
                        
                        # Check again for textarea
                        query_textarea = page.locator("textarea[name='query']")
                        if query_textarea.count() == 0:
                            logger.info("Textarea not found after re-login, trying URL navigation...")
                            
                            # We're now on database page, need to get to SQL query page
                            # Try to navigate via URL with proper params
                            current_url = page.url
                            # Add ?sql= to URL to get SQL page
                            sql_url = current_url + ("&sql=" if "?" in current_url else "?sql=")
                            logger.info(f"Navigating to SQL via URL: {sql_url}")
                            
                            page.goto(sql_url, wait_until='load', timeout=15000)
                            page.wait_for_timeout(1000)  # Extra wait for page render
                            
                            # Check for textarea
                            query_textarea = page.locator("textarea[name='query']")
                            if query_textarea.count() == 0:
                                logger.error(f"Still no textarea after URL navigation. Page title: {page.title()}")
                                logger.error(f"Page URL: {page.url}")
                                
                                # Log all links/buttons for debugging
                                all_links = page.locator("a, button")
                                logger.error(f"Found {all_links.count()} links/buttons on page")
                                for i in range(min(10, all_links.count())):
                                    text = all_links.nth(i).text_content()
                                    href = all_links.nth(i).get_attribute("href") or "N/A"
                                    logger.error(f"  {i}: {text.strip()[:50]} (href: {href})")
                                
                                return False, [], "Cannot find SQL query form - textarea not accessible"
                            else:
                                logger.info("Found textarea after URL navigation!")
                        
                        else:
                            logger.error(f"Page title: {page.title()}")
                            logger.error(f"Page URL: {page.url}")
                            
                            # Log all links/buttons for debugging
                            all_links = page.locator("a, button")
                            logger.error(f"Found {all_links.count()} links/buttons on page")
                            for i in range(min(10, all_links.count())):
                                text = all_links.nth(i).text_content()
                                href = all_links.nth(i).get_attribute("href") or "N/A"
                                logger.error(f"  {i}: {text.strip()[:50]} (href: {href})")
                            
                            return False, [], "Cannot find SQL query form - menu navigation failed"
                    
                    logger.info("Found query textarea, filling query...")
                    query_textarea = page.locator("textarea[name='query']").first
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


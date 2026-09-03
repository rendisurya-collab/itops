"""
Database Connector untuk QueryExecutor - support multiple databases via config.
"""
import json
import logging
from pathlib import Path
from typing import Dict, List, Optional, Tuple
import pymysql

logger = logging.getLogger(__name__)


class DatabaseConnector:
    """
    Database connector yang support multiple databases dari config file.
    """

    def __init__(self, config_file: str):
        """
        Initialize DatabaseConnector.

        Args:
            config_file: Path ke config.json (misal: D:\\mybot\\tools\\coreitops\\bot_core\\config.json)
        """
        self.config_file = Path(config_file)
        self.config: Dict = {}
        self.connections: Dict = {}
        self.databases: Dict = {}

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

    def get_connection(self, db_name: str):
        """
        Get atau create connection ke database tertentu.

        Args:
            db_name: Nama database dari config (misal: 'apis_2022', 'sap_middleware_live')

        Return:
            Connection object atau None jika gagal
        """
        # Check cache
        if db_name in self.connections:
            try:
                conn = self.connections[db_name]
                # Test connection
                conn.ping()
                return conn
            except Exception as e:
                logger.warning(f"Cached connection '{db_name}' dead, reconnecting: {e}")
                del self.connections[db_name]

        # Find database config
        db_config = None
        for db in self.config.get('databases', []):
            if db.get('name') == db_name:
                db_config = db.get('adminer', {})
                break

        if not db_config:
            logger.error(f"Database '{db_name}' tidak ditemukan di config")
            return None

        try:
            # Create connection
            conn = pymysql.connect(
                host=db_config.get('server'),
                user=db_config.get('username'),
                password=db_config.get('password'),
                database=db_config.get('db'),
                charset='utf8mb4',
                cursorclass=pymysql.cursors.DictCursor,
                connect_timeout=10,
            )
            logger.info(f"Connected to database '{db_name}'")
            self.connections[db_name] = conn
            return conn
        except Exception as e:
            logger.error(f"Error connecting to '{db_name}': {e}")
            return None

    def execute_query(self, db_name: str, query: str) -> Tuple[bool, list, str]:
        """
        Execute query ke database tertentu.

        Args:
            db_name: Nama database
            query: SQL query string

        Return:
            (success, rows, error_message)
        """
        conn = self.get_connection(db_name)
        if not conn:
            return False, [], f"Connection ke '{db_name}' gagal"

        try:
            with conn.cursor() as cursor:
                cursor.execute(query)
                rows = cursor.fetchall()
                logger.info(f"Query executed: {len(rows)} rows returned from '{db_name}'")
                return True, rows, None
        except Exception as e:
            error_msg = f"Query execution error: {str(e)}"
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
            # Check apakah query_name match dengan salah satu query file
            for qf in query_files:
                if qf.replace('.sql', '') == query_name or qf == f"{query_name}.sql":
                    return db.get('name')
        return None

    def close_all(self):
        """Close semua connections."""
        for db_name, conn in self.connections.items():
            try:
                conn.close()
                logger.info(f"Closed connection to '{db_name}'")
            except Exception as e:
                logger.error(f"Error closing connection to '{db_name}': {e}")
        self.connections.clear()

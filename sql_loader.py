"""
SQLLoader module untuk auto-scan dan load file .sql dari folder eksternal.
"""
import os
import json
import logging
from pathlib import Path
from typing import Dict

logger = logging.getLogger(__name__)


class SQLLoader:
    """
    Auto-scan folder SQL_FOLDER_PATH dan load semua file .sql ke memori.
    """

    def __init__(self, sql_folder_path: str):
        """
        Initialize SQLLoader.

        Args:
            sql_folder_path: Path ke folder yang berisi file .sql (misal: ./tools/queries atau /app/tools/queries)
        """
        self.sql_folder_path = Path(sql_folder_path)
        self.queries: Dict[str, str] = {}
        self.config: Dict = {}

        # Validasi folder
        if not self.sql_folder_path.exists():
            logger.warning(f"SQL folder tidak ditemukan: {self.sql_folder_path}")
            return

        if not self.sql_folder_path.is_dir():
            logger.error(f"SQL path bukan folder: {self.sql_folder_path}")
            return

        logger.info(f"SQLLoader initialized dengan folder: {self.sql_folder_path}")

    def load_sql_files(self) -> int:
        """
        Scan folder dan load semua file .sql ke memori.

        Return:
            Jumlah file .sql yang berhasil diload
        """
        if not self.sql_folder_path.exists():
            logger.error(f"Folder tidak ditemukan: {self.sql_folder_path}")
            return 0

        self.queries.clear()
        sql_files = list(self.sql_folder_path.glob("*.sql"))

        if not sql_files:
            logger.warning(f"Tidak ada file .sql di {self.sql_folder_path}")
            return 0

        for sql_file in sql_files:
            try:
                with open(sql_file, "r", encoding="utf-8") as f:
                    content = f.read().strip()
                    if content:
                        query_name = sql_file.stem  # nama file tanpa .sql
                        self.queries[query_name] = content
                        logger.info(f"Loaded query: {query_name} ({len(content)} chars)")
                    else:
                        logger.warning(f"File .sql kosong, skip: {sql_file.name}")
            except Exception as e:
                logger.error(f"Error loading {sql_file.name}: {e}")

        logger.info(f"SQLLoader loaded {len(self.queries)} queries dari {self.sql_folder_path}")
        return len(self.queries)

    def load_config(self) -> bool:
        """
        Load query_config.json dari SQL_FOLDER_PATH.

        Return:
            True jika berhasil, False jika gagal
        """
        config_file = self.sql_folder_path / "query_config.json"

        if not config_file.exists():
            logger.warning(f"Config file tidak ditemukan: {config_file}")
            return False

        try:
            with open(config_file, "r", encoding="utf-8") as f:
                self.config = json.load(f)
            logger.info(f"Loaded query config: {len(self.config)} entries")
            return True
        except Exception as e:
            logger.error(f"Error loading config: {e}")
            return False

    def reload_sql_files(self) -> int:
        """
        Reload semua .sql files dari disk tanpa restart bot.

        Return:
            Jumlah file yang di-reload
        """
        logger.info("Reloading SQL files...")
        return self.load_sql_files()

    def reload_config(self) -> bool:
        """
        Reload query_config.json dari disk tanpa restart bot.

        Return:
            True jika berhasil
        """
        logger.info("Reloading query config...")
        return self.load_config()

    def get_query(self, query_name: str) -> str | None:
        """
        Get query by name.

        Args:
            query_name: Nama query (nama file .sql tanpa ekstensi)

        Return:
            Query string, atau None jika tidak ditemukan
        """
        return self.queries.get(query_name)

    def get_all_queries(self) -> Dict[str, str]:
        """Get semua queries."""
        return self.queries.copy()

    def get_config(self) -> Dict:
        """Get semua config."""
        return self.config.copy()

    def get_config_entry(self, query_name: str) -> Dict | None:
        """
        Get config entry untuk query tertentu.

        Args:
            query_name: Nama query

        Return:
            Config dict, atau None jika tidak ditemukan
        """
        return self.config.get(query_name)

    def is_query_enabled(self, query_name: str) -> bool:
        """
        Check apakah query enabled di config.

        Args:
            query_name: Nama query

        Return:
            True jika enabled dan query tersedia, False sebaliknya
        """
        if query_name not in self.queries:
            return False

        entry = self.config.get(query_name, {})
        return entry.get("enabled", False)

    def list_enabled_queries(self) -> list:
        """
        Get list query_name yang enabled di config dan tersedia di disk.

        Return:
            List of query names yang enabled
        """
        enabled = []
        for query_name in self.queries.keys():
            if self.is_query_enabled(query_name):
                enabled.append(query_name)
        return enabled

    def validate_query_config(self, query_name: str) -> tuple[bool, str]:
        """
        Validasi entry config untuk query tertentu.

        Args:
            query_name: Nama query

        Return:
            (is_valid, error_message)
        """
        if query_name not in self.queries:
            return False, f"Query '{query_name}' tidak ditemukan di disk"

        entry = self.config.get(query_name)
        if not entry:
            return False, f"Config entry untuk '{query_name}' tidak ditemukan"

        # Check required fields
        schedule_type = entry.get("schedule_type")
        if schedule_type not in ["cron", "interval"]:
            return False, f"schedule_type harus 'cron' atau 'interval', diterima: {schedule_type}"

        if schedule_type == "cron":
            if "hour" not in entry or "minute" not in entry:
                return False, "Cron schedule memerlukan 'hour' dan 'minute'"
        elif schedule_type == "interval":
            if "interval_minutes" not in entry:
                return False, "Interval schedule memerlukan 'interval_minutes'"

        return True, ""

"""
Dynamic Scheduler module untuk auto-register dan execute query jobs dengan APScheduler.
"""
import logging
import asyncio
import datetime as dt
from typing import Dict, Callable, Optional

from apscheduler.schedulers.asyncio import AsyncIOScheduler
from apscheduler.triggers.cron import CronTrigger
from apscheduler.triggers.interval import IntervalTrigger
from apscheduler.job import Job

from sql_loader import SQLLoader
from query_executor import QueryExecutor

logger = logging.getLogger(__name__)


class DynamicScheduler:
    """
    Auto-register dan execute SQL queries berdasarkan config dinamis.
    """

    def __init__(
        self,
        sql_loader: SQLLoader,
        query_executor: QueryExecutor,
        telegram_notify_callback: Optional[Callable] = None,
    ):
        """
        Initialize DynamicScheduler.

        Args:
            sql_loader: Instance SQLLoader
            query_executor: Instance QueryExecutor
            telegram_notify_callback: Async callback untuk send Telegram notification
                                     signature: async def callback(chat_id, text, excel_bytes=None)
        """
        self.sql_loader = sql_loader
        self.query_executor = query_executor
        self.telegram_notify_callback = telegram_notify_callback
        self.scheduler: AsyncIOScheduler = None
        self.registered_jobs: Dict[str, Job] = {}

    def initialize_scheduler(self) -> bool:
        """
        Initialize APScheduler (AsyncIOScheduler).

        Return:
            True jika berhasil
        """
        try:
            self.scheduler = AsyncIOScheduler()
            logger.info("AsyncIOScheduler initialized")
            return True
        except Exception as e:
            logger.error(f"Error initializing scheduler: {e}")
            return False

    def start_scheduler(self) -> bool:
        """
        Start scheduler background job.

        Return:
            True jika berhasil
        """
        if not self.scheduler:
            logger.error("Scheduler belum initialized")
            return False

        try:
            if not self.scheduler.running:
                self.scheduler.start()
                logger.info("Scheduler started")
            return True
        except Exception as e:
            logger.error(f"Error starting scheduler: {e}")
            return False

    def stop_scheduler(self) -> bool:
        """Stop scheduler gracefully."""
        if not self.scheduler:
            return True

        try:
            if self.scheduler.running:
                self.scheduler.shutdown(wait=False)
                logger.info("Scheduler stopped")
            return True
        except Exception as e:
            logger.error(f"Error stopping scheduler: {e}")
            return False

    async def _execute_query_job(self, query_name: str, chat_id: str | None = None):
        """
        Execute query job dan kirim notifikasi ke Telegram.

        Workflow:
        1. Get query dari sql_loader
        2. Execute query
        3. Jika no rows -> log saja, tidak kirim notif
        4. Jika ada rows -> process hasil dan kirim notif (teks atau Excel)

        Args:
            query_name: Nama query
            chat_id: Telegram chat ID untuk notifikasi (optional, ambil dari config kalau None)
        """
        try:
            # Get query
            query = self.sql_loader.get_query(query_name)
            if not query:
                logger.warning(f"Query '{query_name}' tidak ditemukan")
                return

            logger.info(f"Executing job: {query_name}")

            # Get chat_id dari config jika tidak diberikan
            if not chat_id:
                config_entry = self.sql_loader.get_config_entry(query_name)
                if config_entry:
                    chat_id = config_entry.get("telegram_chat_id", "").strip()

            # Fallback to default NOTIFY_TARGETS jika chat_id masih kosong
            if not chat_id:
                from config import notify_targets
                targets = notify_targets()
                if targets:
                    # Ambil target pertama (chat_id, thread_id)
                    chat_id, thread_id = targets[0]
                else:
                    logger.warning(f"Tidak ada chat_id di config dan NOTIFY_TARGETS kosong")
                    return

            # Execute query di background thread
            success, rows, error = await asyncio.to_thread(
                self.query_executor.execute_select, query
            )

            if not success:
                logger.error(f"Query execution failed: {error}")
                return

            # Check if empty
            if not rows:
                logger.info(f"Query '{query_name}' returned 0 rows, skip notification")
                return

            logger.info(f"Query '{query_name}' returned {len(rows)} rows")

            # Process result (text + optional Excel)
            text_message, excel_bytes, process_error = await asyncio.to_thread(
                self.query_executor.process_query_result,
                rows,
                query_name,
                3500
            )

            # Send notification
            if self.telegram_notify_callback and chat_id:
                try:
                    await self.telegram_notify_callback(
                        chat_id=chat_id,
                        text=f"Query: {query_name}\n\n{text_message}",
                        excel_bytes=excel_bytes,
                        filename=f"{query_name}_{dt.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                    )
                    logger.info(f"Notification sent untuk job: {query_name}")
                except Exception as e:
                    logger.error(f"Error sending notification: {e}")
            else:
                logger.warning(f"Telegram callback tidak tersedia atau chat_id kosong")

        except Exception as e:
            logger.exception(f"Error executing job '{query_name}': {e}")

    def register_jobs_from_config(self) -> int:
        """
        Auto-register jobs dari query_config.json.

        Return:
            Jumlah job yang berhasil didaftarkan
        """
        if not self.scheduler:
            logger.error("Scheduler belum initialized")
            return 0

        enabled_queries = self.sql_loader.list_enabled_queries()

        if not enabled_queries:
            logger.warning("Tidak ada query yang enabled")
            return 0

        registered_count = 0

        for query_name in enabled_queries:
            # Validate config
            is_valid, error_msg = self.sql_loader.validate_query_config(query_name)
            if not is_valid:
                logger.error(f"Invalid config untuk '{query_name}': {error_msg}")
                continue

            # Get config entry
            config_entry = self.sql_loader.get_config_entry(query_name)
            schedule_type = config_entry.get("schedule_type")

            try:
                if schedule_type == "cron":
                    # Register cron job
                    hour = config_entry.get("hour", 0)
                    minute = config_entry.get("minute", 0)
                    description = config_entry.get("description", f"Cron job: {query_name}")

                    trigger = CronTrigger(hour=hour, minute=minute)
                    job = self.scheduler.add_job(
                        self._execute_query_job,
                        trigger=trigger,
                        args=[query_name],
                        name=f"cron_{query_name}",
                        id=f"cron_{query_name}",
                        replace_existing=True,
                    )
                    logger.info(f"Registered cron job: {query_name} at {hour:02d}:{minute:02d}")
                    registered_count += 1

                elif schedule_type == "interval":
                    # Register interval job
                    interval_minutes = config_entry.get("interval_minutes", 60)
                    description = config_entry.get("description", f"Interval job: {query_name}")

                    trigger = IntervalTrigger(minutes=interval_minutes)
                    job = self.scheduler.add_job(
                        self._execute_query_job,
                        trigger=trigger,
                        args=[query_name],
                        name=f"interval_{query_name}",
                        id=f"interval_{query_name}",
                        replace_existing=True,
                    )
                    logger.info(f"Registered interval job: {query_name} every {interval_minutes} minutes")
                    registered_count += 1

            except Exception as e:
                logger.error(f"Error registering job for '{query_name}': {e}")

        logger.info(f"Registered {registered_count}/{len(enabled_queries)} jobs")
        return registered_count

    def unregister_job(self, query_name: str) -> bool:
        """
        Unregister job by query name.

        Args:
            query_name: Nama query

        Return:
            True jika berhasil
        """
        if not self.scheduler:
            return False

        for job_id in [f"cron_{query_name}", f"interval_{query_name}"]:
            try:
                job = self.scheduler.get_job(job_id)
                if job:
                    self.scheduler.remove_job(job_id)
                    logger.info(f"Unregistered job: {job_id}")
                    return True
            except:
                pass

        return False

    def get_registered_jobs(self) -> list:
        """Get list semua registered jobs."""
        if not self.scheduler:
            return []

        return [
            {
                "name": job.name,
                "id": job.id,
                "trigger": str(job.trigger),
                "next_run": getattr(job, 'next_run_time', None),
            }
            for job in self.scheduler.get_jobs()
        ]

    def reload_and_reregister(self) -> int:
        """
        Reload config & SQL files, lalu re-register semua jobs.

        Return:
            Jumlah job yang di-re-register
        """
        logger.info("Reloading SQL loader...")

        # Reload files dan config
        self.sql_loader.reload_sql_files()
        self.sql_loader.reload_config()

        # Remove existing jobs
        if self.scheduler:
            for job in self.scheduler.get_jobs():
                self.scheduler.remove_job(job.id)
            logger.info("Cleared all existing jobs")

        # Re-register
        return self.register_jobs_from_config()

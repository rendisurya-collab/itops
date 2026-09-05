"""
Reminders Scheduler - Register and execute APScheduler jobs for reminders.
Supports both interval-based and cron-based (scheduled) reminders.
"""

import logging
from typing import Optional, Dict, Callable
from apscheduler.schedulers.asyncio import AsyncIOScheduler
from apscheduler.triggers.interval import IntervalTrigger
from apscheduler.triggers.cron import CronTrigger
from telegram.ext import Application, ContextTypes
from reminders_manager import RemindersManager, ReminderConfig

logger = logging.getLogger(__name__)


class RemindersScheduler:
    """Manages APScheduler jobs for all reminders."""
    
    def __init__(self, application: Application, reminders_manager: RemindersManager):
        """
        Initialize RemindersScheduler.
        
        Args:
            application: telegram.ext.Application instance
            reminders_manager: RemindersManager instance
        """
        self.application = application
        self.manager = reminders_manager
        self.job_queue = application.job_queue
        
        # Get the actual APScheduler scheduler from job_queue
        # The scheduler is stored as _scheduler in modern python-telegram-bot
        try:
            self.scheduler = getattr(self.job_queue, '_scheduler', None)
            if self.scheduler is None:
                # Try alternate access path
                self.scheduler = getattr(self.job_queue, 'scheduler', None)
        except:
            self.scheduler = None
        
        if self.scheduler is None:
            logger.warning("⚠️ Could not access APScheduler directly, will use job_queue methods")
        
        self.active_jobs: Dict[str, str] = {}  # reminder_id -> job_id mapping
    
    async def send_reminder_notification(self, reminder: ReminderConfig, context: ContextTypes.DEFAULT_TYPE) -> None:
        """
        Send reminder notification to all configured chat targets.
        
        Args:
            reminder: ReminderConfig instance
            context: Telegram context
        """
        try:
            from config import notify_targets
            from datetime import datetime
            import pytz
            
            # Check time range for interval reminders
            if reminder.interval_type == "interval" and hasattr(reminder, 'hour_start') and reminder.hour_start is not None:
                tz = pytz.timezone('Asia/Jakarta')
                current_hour = datetime.now(tz).hour
                
                if not (reminder.hour_start <= current_hour < reminder.hour_end):
                    logger.info(f"⊘ Reminder '{reminder.id}' skipped (outside time range {reminder.hour_start:02d}:00-{reminder.hour_end:02d}:00, current: {current_hour:02d}:00)")
                    return
            
            targets = notify_targets()
            if not targets:
                logger.warning(f"⚠️ No notification targets configured for reminder {reminder.id}")
                return
            
            for chat_id, thread_id in targets:
                try:
                    await context.bot.send_message(
                        chat_id=chat_id,
                        text=reminder.message,
                        parse_mode="HTML",
                        message_thread_id=thread_id
                    )
                    logger.info(f"✓ Reminder '{reminder.id}' sent to {chat_id}")
                except Exception as e:
                    logger.error(f"❌ Failed to send reminder '{reminder.id}' to {chat_id}: {e}")
        
        except Exception as e:
            logger.error(f"❌ Error in send_reminder_notification: {e}", exc_info=True)
    
    async def _reminder_job_wrapper(self, reminder_id: str) -> None:
        """
        Wrapper function for reminder job (required for async execution).
        
        Args:
            reminder_id: ID of reminder to execute
        """
        try:
            reminder = self.manager.get_reminder_by_id(reminder_id)
            if not reminder:
                logger.warning(f"⚠️ Reminder not found: {reminder_id}")
                return
            
            # Get context from application
            context = ContextTypes.DEFAULT_TYPE()
            context.bot = self.application.bot
            context.application = self.application
            
            await self.send_reminder_notification(reminder, context)
        
        except Exception as e:
            logger.error(f"❌ Error executing reminder job '{reminder_id}': {e}", exc_info=True)
    
    def _reminder_job_wrapper_sync(self, reminder_id: str, context=None) -> None:
        """
        Sync wrapper for reminder job (for job_queue fallback).
        
        Args:
            reminder_id: ID of reminder to execute
            context: Telegram context (from job_queue)
        """
        try:
            import asyncio
            
            reminder = self.manager.get_reminder_by_id(reminder_id)
            if not reminder:
                logger.warning(f"⚠️ Reminder not found: {reminder_id}")
                return
            
            # Get or create context
            if context is None:
                ctx = ContextTypes.DEFAULT_TYPE()
                ctx.bot = self.application.bot
                ctx.application = self.application
            else:
                ctx = context
            
            # Run async function in event loop
            loop = asyncio.new_event_loop()
            asyncio.set_event_loop(loop)
            try:
                loop.run_until_complete(self.send_reminder_notification(reminder, ctx))
            finally:
                loop.close()
        
        except Exception as e:
            logger.error(f"❌ Error executing sync reminder job '{reminder_id}': {e}", exc_info=True)
    
    def _create_interval_trigger(self, reminder: ReminderConfig) -> IntervalTrigger:
        """Create IntervalTrigger for interval-based reminder."""
        return IntervalTrigger(minutes=reminder.interval_minutes)
    
    def _create_cron_trigger(self, reminder: ReminderConfig) -> CronTrigger:
        """Create CronTrigger for cron-based reminder."""
        trigger_kwargs = {
            'hour': reminder.hour,
            'minute': reminder.minute,
            'timezone': 'Asia/Jakarta'
        }
        
        # Add day_of_week if specified
        if hasattr(reminder, 'day_of_week') and reminder.day_of_week:
            trigger_kwargs['day_of_week'] = reminder.day_of_week
        
        # Add day_of_month if specified
        if hasattr(reminder, 'day_of_month') and reminder.day_of_month:
            trigger_kwargs['day'] = reminder.day_of_month
        
        return CronTrigger(**trigger_kwargs)
    
    def register_reminder(self, reminder: ReminderConfig) -> bool:
        """
        Register a single reminder with scheduler.
        
        Args:
            reminder: ReminderConfig instance
            
        Return:
            bool: True if registered successfully, False if error
        """
        try:
            if not reminder.enabled:
                logger.info(f"⊘ Reminder '{reminder.id}' is disabled, skipping")
                return True
            
            # Create trigger based on interval_type
            if reminder.interval_type == "interval":
                trigger = self._create_interval_trigger(reminder)
                logger.info(f"Creating interval trigger for {reminder.id}: every {reminder.interval_minutes} minutes")
                
                # Use job_queue.run_repeating for interval reminders
                if self.scheduler:
                    # Remove existing job if any
                    if reminder.id in self.active_jobs:
                        try:
                            self.scheduler.remove_job(self.active_jobs[reminder.id])
                        except Exception as e:
                            logger.warning(f"Could not remove old job for {reminder.id}: {e}")
                        del self.active_jobs[reminder.id]
                    
                    job = self.scheduler.add_job(
                        self._reminder_job_wrapper,
                        trigger=trigger,
                        args=[reminder.id],
                        id=f"reminder_{reminder.id}",
                        name=reminder.name,
                        replace_existing=True,
                        max_instances=1
                    )
                    self.active_jobs[reminder.id] = job.id
                else:
                    # Fallback: use job_queue.run_repeating
                    logger.info(f"Using job_queue.run_repeating for {reminder.id}")
                    import datetime as dt
                    self.job_queue.run_repeating(
                        lambda ctx: self._reminder_job_wrapper_sync(reminder.id),
                        interval=dt.timedelta(minutes=reminder.interval_minutes),
                        first=1,
                        name=f"reminder_{reminder.id}"
                    )
                    self.active_jobs[reminder.id] = f"reminder_{reminder.id}"
            
            else:  # cron
                trigger = self._create_cron_trigger(reminder)
                logger.info(f"Creating cron trigger for {reminder.id}")
                
                if self.scheduler:
                    # Remove existing job if any
                    if reminder.id in self.active_jobs:
                        try:
                            self.scheduler.remove_job(self.active_jobs[reminder.id])
                        except Exception as e:
                            logger.warning(f"Could not remove old job for {reminder.id}: {e}")
                        del self.active_jobs[reminder.id]
                    
                    # Register job with scheduler
                    job = self.scheduler.add_job(
                        self._reminder_job_wrapper,
                        trigger=trigger,
                        args=[reminder.id],
                        id=f"reminder_{reminder.id}",
                        name=reminder.name,
                        replace_existing=True,
                        max_instances=1
                    )
                    self.active_jobs[reminder.id] = job.id
                else:
                    # Fallback: use job_queue.run_daily or similar
                    logger.info(f"Using job_queue.run_daily for {reminder.id}")
                    import datetime as dt
                    self.job_queue.run_daily(
                        lambda ctx: self._reminder_job_wrapper_sync(reminder.id),
                        time=dt.time(hour=reminder.hour, minute=reminder.minute),
                        name=f"reminder_{reminder.id}"
                    )
                    self.active_jobs[reminder.id] = f"reminder_{reminder.id}"
            
            logger.info(f"✓ Registered reminder '{reminder.name}' ({reminder.id})")
            return True
        
        except Exception as e:
            logger.error(f"❌ Error registering reminder '{reminder.id}': {e}", exc_info=True)
            return False
    
    def register_all_reminders(self) -> tuple[int, int]:
        """
        Register all active reminders from manager.
        
        Return:
            (registered_count, failed_count)
        """
        reminders = self.manager.get_active_reminders()
        
        registered = 0
        failed = 0
        
        for reminder in reminders:
            if self.register_reminder(reminder):
                registered += 1
            else:
                failed += 1
        
        logger.info(f"✓✓ Reminders registration complete: {registered} registered, {failed} failed")
        return registered, failed
    
    def unregister_reminder(self, reminder_id: str) -> bool:
        """
        Unregister (remove) a reminder from scheduler.
        
        Args:
            reminder_id: ID of reminder to remove
            
        Return:
            bool: True if removed successfully
        """
        try:
            if reminder_id not in self.active_jobs:
                logger.warning(f"⚠️ Reminder '{reminder_id}' is not registered")
                return False
            
            job_id = self.active_jobs[reminder_id]
            
            if self.scheduler:
                self.scheduler.remove_job(job_id)
            else:
                logger.info(f"Job {job_id} not removed (scheduler not accessible)")
            
            del self.active_jobs[reminder_id]
            
            logger.info(f"✓ Unregistered reminder '{reminder_id}'")
            return True
        
        except Exception as e:
            logger.error(f"❌ Error unregistering reminder '{reminder_id}': {e}")
            return False
    
    def unregister_all_reminders(self) -> int:
        """
        Unregister all reminders from scheduler.
        
        Return:
            int: Number of reminders unregistered
        """
        reminder_ids = list(self.active_jobs.keys())
        count = 0
        
        for reminder_id in reminder_ids:
            if self.unregister_reminder(reminder_id):
                count += 1
        
        logger.info(f"✓ Unregistered {count} reminders")
        return count
    
    async def reload_all_reminders(self) -> tuple[bool, str]:
        """
        Reload configuration and re-register all reminders (hot-reload).
        
        Return:
            (success: bool, message: str)
        """
        try:
            # Reload config from file
            success, msg = self.manager.reload_config()
            if not success:
                return False, msg
            
            # Unregister all old jobs
            unregistered = self.unregister_all_reminders()
            
            # Register new reminders
            registered, failed = self.register_all_reminders()
            
            if failed > 0:
                return False, f"⚠️ Reloaded with {registered} registered, {failed} failed"
            else:
                return True, f"✓ Reminders reloaded: {registered} registered, {unregistered} unregistered"
        
        except Exception as e:
            msg = f"❌ Error reloading reminders: {e}"
            logger.error(msg, exc_info=True)
            return False, msg
    
    def get_status(self) -> str:
        """Get current scheduler status."""
        lines = []
        lines.append("📊 <b>REMINDERS SCHEDULER STATUS</b>\n")
        lines.append(f"Active jobs: {len(self.active_jobs)}")
        
        if self.scheduler:
            lines.append(f"Scheduler running: {self.scheduler.running}")
        else:
            lines.append("Scheduler: Using job_queue (direct access not available)")
        
        if self.active_jobs:
            lines.append("\n<b>Registered reminders:</b>")
            for reminder_id, job_id in self.active_jobs.items():
                reminder = self.manager.get_reminder_by_id(reminder_id)
                if reminder:
                    lines.append(f"  • {reminder.name} ({reminder_id})")
        
        return "\n".join(lines)


# Global scheduler instance
_scheduler: Optional[RemindersScheduler] = None


async def init_reminders_scheduler(application: Application, reminders_manager: RemindersManager) -> RemindersScheduler:
    """Initialize global reminders scheduler."""
    global _scheduler
    _scheduler = RemindersScheduler(application, reminders_manager)
    
    # Register all reminders
    registered, failed = _scheduler.register_all_reminders()
    if failed > 0:
        logger.warning(f"⚠️ Some reminders failed to register ({failed} failed)")
    
    return _scheduler


def get_reminders_scheduler() -> Optional[RemindersScheduler]:
    """Get global reminders scheduler."""
    global _scheduler
    return _scheduler

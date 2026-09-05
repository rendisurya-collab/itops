"""
Reminders Manager - Load and manage reminder configurations from JSON file.
Supports hot-reload for dynamic configuration updates without restarting bot.
"""

import json
import logging
from pathlib import Path
from typing import List, Dict, Optional
from datetime import datetime

logger = logging.getLogger(__name__)

DEFAULT_CONFIG_FILE = "reminders_config.json"


class ReminderConfig:
    """Represents a single reminder configuration."""
    
    def __init__(self, config_dict: dict):
        """
        Initialize reminder config from dict.
        
        Args:
            config_dict: Dictionary with reminder configuration
            
        Raises:
            ValueError: If configuration is invalid
        """
        self.id = config_dict.get("id")
        self.name = config_dict.get("name")
        self.enabled = config_dict.get("enabled", True)
        self.interval_type = config_dict.get("interval_type")  # "interval" or "cron"
        self.message = config_dict.get("message")
        
        # Validate required fields
        if not self.id:
            raise ValueError("Reminder must have 'id' field")
        if not self.name:
            raise ValueError(f"Reminder {self.id} must have 'name' field")
        if not self.interval_type:
            raise ValueError(f"Reminder {self.id} must have 'interval_type' field")
        if self.interval_type not in ["interval", "cron"]:
            raise ValueError(f"Reminder {self.id}: interval_type must be 'interval' or 'cron', got {self.interval_type}")
        if not self.message:
            raise ValueError(f"Reminder {self.id} must have 'message' field")
        
        # Interval-specific config
        if self.interval_type == "interval":
            self.interval_minutes = config_dict.get("interval_minutes", 60)
            if not isinstance(self.interval_minutes, int) or self.interval_minutes <= 0:
                raise ValueError(f"Reminder {self.id}: interval_minutes must be positive integer")
        
        # Cron-specific config
        if self.interval_type == "cron":
            self.hour = config_dict.get("hour", 0)
            self.minute = config_dict.get("minute", 0)
            self.day_of_week = config_dict.get("day_of_week")  # "mon", "tue", etc. or None
            self.day_of_month = config_dict.get("day_of_month")  # 1-31 or None
            
            # Validate hour/minute
            if not isinstance(self.hour, int) or not (0 <= self.hour < 24):
                raise ValueError(f"Reminder {self.id}: hour must be 0-23")
            if not isinstance(self.minute, int) or not (0 <= self.minute < 60):
                raise ValueError(f"Reminder {self.id}: minute must be 0-59")
            
            # Must have either day_of_week or day_of_month
            if not self.day_of_week and not self.day_of_month:
                raise ValueError(f"Reminder {self.id}: cron reminder must have day_of_week or day_of_month")
    
    def __repr__(self) -> str:
        status = "✓" if self.enabled else "✗"
        return f"{status} {self.name} ({self.id}) - {self.interval_type}"


class RemindersManager:
    """Manages loading and reloading reminder configurations."""
    
    def __init__(self, config_file: str = DEFAULT_CONFIG_FILE):
        """
        Initialize RemindersManager.
        
        Args:
            config_file: Path to reminders_config.json
        """
        self.config_file = Path(config_file)
        self.reminders: List[ReminderConfig] = []
        self._last_load_time: Optional[datetime] = None
        self.load_config()
    
    def load_config(self) -> bool:
        """
        Load reminder configurations from JSON file.
        
        Return:
            bool: True if loaded successfully, False if error
        """
        try:
            if not self.config_file.exists():
                logger.error(f"❌ Reminders config file not found: {self.config_file}")
                return False
            
            with open(self.config_file, 'r', encoding='utf-8') as f:
                data = json.load(f)
            
            if not isinstance(data, dict) or "reminders" not in data:
                logger.error("❌ Invalid reminders config format: must have 'reminders' array")
                return False
            
            reminders_data = data.get("reminders", [])
            if not isinstance(reminders_data, list):
                logger.error("❌ Invalid reminders config: 'reminders' must be array")
                return False
            
            # Parse all reminders
            new_reminders = []
            for idx, reminder_dict in enumerate(reminders_data):
                try:
                    reminder = ReminderConfig(reminder_dict)
                    new_reminders.append(reminder)
                except ValueError as e:
                    logger.error(f"❌ Error parsing reminder #{idx}: {e}")
                    return False
            
            self.reminders = new_reminders
            self._last_load_time = datetime.now()
            
            logger.info(f"✓ Loaded {len(self.reminders)} reminders from {self.config_file}")
            for reminder in self.reminders:
                logger.info(f"  - {reminder}")
            
            return True
            
        except json.JSONDecodeError as e:
            logger.error(f"❌ JSON parse error in reminders config: {e}")
            return False
        except Exception as e:
            logger.error(f"❌ Error loading reminders config: {e}", exc_info=True)
            return False
    
    def get_active_reminders(self) -> List[ReminderConfig]:
        """Get list of enabled reminders only."""
        return [r for r in self.reminders if r.enabled]
    
    def get_reminder_by_id(self, reminder_id: str) -> Optional[ReminderConfig]:
        """Get reminder by ID."""
        for reminder in self.reminders:
            if reminder.id == reminder_id:
                return reminder
        return None
    
    def reload_config(self) -> tuple[bool, str]:
        """
        Reload configuration (hot-reload).
        
        Return:
            (success: bool, message: str)
        """
        try:
            old_count = len(self.reminders)
            if self.load_config():
                new_count = len(self.reminders)
                active_count = len(self.get_active_reminders())
                msg = f"✓ Reminders config reloaded: {new_count} total, {active_count} active (was {old_count})"
                logger.info(msg)
                return True, msg
            else:
                return False, "❌ Failed to reload reminders config"
        except Exception as e:
            msg = f"❌ Error reloading reminders config: {e}"
            logger.error(msg)
            return False, msg
    
    def get_summary(self) -> str:
        """Get summary of all reminders for display."""
        if not self.reminders:
            return "❌ No reminders configured"
        
        lines = []
        lines.append("📋 **DAFTAR REMINDERS**\n")
        
        active = self.get_active_reminders()
        inactive = [r for r in self.reminders if not r.enabled]
        
        if active:
            lines.append(f"✓ <b>ACTIVE ({len(active)}):</b>")
            for reminder in active:
                lines.append(f"  • {reminder.name} ({reminder.id})")
                if reminder.interval_type == "interval":
                    lines.append(f"    └─ Every {reminder.interval_minutes} minutes")
                else:
                    schedule_info = []
                    if hasattr(reminder, 'day_of_week') and reminder.day_of_week:
                        schedule_info.append(f"{reminder.day_of_week}")
                    if hasattr(reminder, 'day_of_month') and reminder.day_of_month:
                        schedule_info.append(f"day {reminder.day_of_month}")
                    lines.append(f"    └─ {', '.join(schedule_info)} @ {reminder.hour:02d}:{reminder.minute:02d}")
            lines.append("")
        
        if inactive:
            lines.append(f"✗ <b>INACTIVE ({len(inactive)}):</b>")
            for reminder in inactive:
                lines.append(f"  • {reminder.name} ({reminder.id})")
            lines.append("")
        
        return "\n".join(lines)


# Global manager instance
_manager: Optional[RemindersManager] = None


def init_reminders_manager(config_file: str = DEFAULT_CONFIG_FILE) -> RemindersManager:
    """Initialize global reminders manager."""
    global _manager
    _manager = RemindersManager(config_file)
    return _manager


def get_reminders_manager() -> RemindersManager:
    """Get global reminders manager."""
    global _manager
    if _manager is None:
        _manager = RemindersManager()
    return _manager

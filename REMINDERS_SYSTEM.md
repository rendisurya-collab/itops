# Dynamic Reminders System Documentation

## Overview

The Dynamic Reminders System is a configurable, hot-reload-capable reminder management system for Telegram Bot. It replaces hardcoded reminder functions with a flexible JSON-based configuration, allowing administrators to add, modify, or remove reminders without changing code or restarting the bot.

**Key Features:**
- ✓ JSON-based configuration (reminders_config.json)
- ✓ Support for interval-based reminders (every X minutes)
- ✓ Support for cron-based reminders (scheduled times/days)
- ✓ Hot-reload capability via `/remindersreload` command
- ✓ Flexible notification routing to all configured chat targets
- ✓ Full error handling and validation
- ✓ Status monitoring via `/reminderslist` command

---

## Architecture

### Components

#### 1. **reminders_config.json**
Central configuration file defining all reminders.

```json
{
  "reminders": [
    {
      "id": "reminder_logwork",
      "name": "Reminder: Logwork",
      "enabled": true,
      "interval_type": "interval",
      "interval_minutes": 360,
      "message": "⏰ <b>Reminder Logwork:</b> Jangan lupa isi logwork guys"
    },
    {
      "id": "reminder_absensi_bulanan",
      "name": "Reminder: Absensi Bulanan",
      "enabled": true,
      "interval_type": "cron",
      "day_of_month": 25,
      "hour": 9,
      "minute": 0,
      "message": "📅 <b>Reminder Absensi:</b> Mohon rekap dan submit absensi bulanan"
    },
    {
      "id": "reminder_weekly_logwork",
      "name": "Reminder: Weekly Logwork",
      "enabled": true,
      "interval_type": "cron",
      "day_of_week": "fri",
      "hour": 17,
      "minute": 0,
      "message": "📝 <b>Reminder Weekly:</b> Jangan lupa submit catatan log mingguan!"
    }
  ]
}
```

#### 2. **reminders_manager.py**
Loads and manages reminder configurations with validation.

**Key Classes:**
- `ReminderConfig`: Represents a single reminder with validation
- `RemindersManager`: Manages loading, validating, and hot-reloading configs

**Key Methods:**
```python
manager = init_reminders_manager("reminders_config.json")
manager.get_active_reminders()           # Get enabled reminders only
manager.get_reminder_by_id(id)          # Find reminder by ID
manager.reload_config()                 # Reload from file
manager.get_summary()                   # Get formatted display
```

#### 3. **reminders_scheduler.py**
Registers and executes APScheduler jobs for reminders.

**Key Classes:**
- `RemindersScheduler`: Wraps APScheduler, manages job registration/execution

**Key Methods:**
```python
scheduler = await init_reminders_scheduler(app, manager)
scheduler.register_reminder(reminder)           # Register single reminder
scheduler.register_all_reminders()             # Register all enabled
scheduler.unregister_reminder(id)              # Remove a reminder
scheduler.reload_all_reminders()               # Hot-reload (unregister + re-register)
scheduler.get_status()                        # Get formatted status
```

#### 4. **bot.py Integration**
- Initializes reminders manager on startup
- Lazily initializes scheduler via `post_init` handler
- Provides `/reminderslist` and `/remindersreload` commands

---

## Configuration Guide

### reminders_config.json Format

Each reminder entry requires:

| Field | Type | Required | Description |
|-------|------|----------|-------------|
| `id` | string | Yes | Unique identifier for reminder |
| `name` | string | Yes | Display name |
| `enabled` | boolean | No | Enable/disable without deleting (default: true) |
| `interval_type` | string | Yes | Either `"interval"` or `"cron"` |
| `message` | string | Yes | Telegram message text (supports HTML) |

**For `interval_type: "interval"`:**
| Field | Type | Required | Description |
|-------|------|----------|-------------|
| `interval_minutes` | integer | Yes | Repeat every N minutes (must be > 0) |

**For `interval_type: "cron"`:**
| Field | Type | Required | Description |
|-------|------|----------|-------------|
| `hour` | integer | Yes | Hour (0-23) |
| `minute` | integer | Yes | Minute (0-59) |
| `day_of_week` | string | No | Day name: "mon", "tue", "wed", "thu", "fri", "sat", "sun" |
| `day_of_month` | integer | No | Day of month (1-31) |

**Note:** For cron reminders, must specify EITHER `day_of_week` OR `day_of_month` (or both)

### Example Configurations

#### Every 6 hours (Interval)
```json
{
  "id": "reminder_logwork_6h",
  "name": "Logwork Reminder (6h)",
  "enabled": true,
  "interval_type": "interval",
  "interval_minutes": 360,
  "message": "⏰ Reminder: jangan lupa isi logwork"
}
```

#### Monthly on 25th at 9:00 AM
```json
{
  "id": "reminder_monthly_absensi",
  "name": "Monthly Absensi",
  "enabled": true,
  "interval_type": "cron",
  "day_of_month": 25,
  "hour": 9,
  "minute": 0,
  "message": "📅 Reminder: Isi absensi bulanan"
}
```

#### Every Friday at 5:00 PM
```json
{
  "id": "reminder_friday_standup",
  "name": "Friday Standup",
  "enabled": true,
  "interval_type": "cron",
  "day_of_week": "fri",
  "hour": 17,
  "minute": 0,
  "message": "📝 Reminder: Friday standup meeting at 17:30"
}
```

#### Combine: Every Monday at 8:00 AM + 15th & 25th
```json
{
  "id": "reminder_weekstart",
  "name": "Week Start Reminder",
  "enabled": true,
  "interval_type": "cron",
  "day_of_week": "mon",
  "day_of_month": 15,
  "hour": 8,
  "minute": 0,
  "message": "🔔 Week started! Check your tasks"
}
```
*Note: This runs on Mondays AND on the 15th of each month*

---

## Commands

### `/reminderslist`
Displays active and inactive reminders with scheduler status.

**Output:**
```
📋 DAFTAR REMINDERS

✓ ACTIVE (3):
  • Reminder: Logwork (reminder_logwork)
    └─ Every 360 minutes
  • Reminder: Absensi Bulanan (reminder_absensi_bulanan)
    └─ day 25 @ 09:00
  • Reminder: Submit Catatan Log (reminder_submit_catatan_log)
    └─ fri @ 17:00

📊 REMINDERS SCHEDULER STATUS

Active jobs: 3
Scheduler running: True

Registered reminders:
  • Reminder: Logwork (reminder_logwork)
  • Reminder: Absensi Bulanan (reminder_absensi_bulanan)
  • Reminder: Submit Catatan Log (reminder_submit_catatan_log)
```

### `/remindersreload`
Reloads reminders_config.json and re-registers all jobs (hot-reload).

Useful when:
- Adding new reminders to reminders_config.json
- Disabling/enabling existing reminders
- Changing reminder schedules or messages

**No bot restart required!**

---

## Usage Workflows

### Adding a New Reminder

1. Edit `reminders_config.json`
2. Add new reminder entry to `reminders` array:
   ```json
   {
     "id": "reminder_new_feature",
     "name": "New Feature Reminder",
     "enabled": true,
     "interval_type": "cron",
     "day_of_week": "wed",
     "hour": 14,
     "minute": 0,
     "message": "🎉 New feature reminder!"
   }
   ```
3. Send `/remindersreload` in Telegram
4. Verify with `/reminderslist`

### Temporarily Disabling a Reminder

1. Edit `reminders_config.json`
2. Set `"enabled": false` for the reminder
3. Send `/remindersreload`

### Modifying a Reminder Schedule

1. Edit `reminders_config.json`
2. Change `interval_minutes` (for interval) or `hour`/`minute`/`day_of_week`/`day_of_month` (for cron)
3. Send `/remindersreload`

### Deleting a Reminder

1. Edit `reminders_config.json`
2. Remove the entire reminder entry from `reminders` array
3. Send `/remindersreload`

---

## Technical Details

### Trigger Types

#### IntervalTrigger
- Repeats every N minutes
- Simpler than cron, useful for regular intervals
- Example: "Every 6 hours" (360 minutes)

#### CronTrigger
- Scheduled at specific times/days
- More complex but flexible
- Timezone: Asia/Jakarta (configurable in reminders_scheduler.py)

### Notification Routing

Reminders are sent to all targets defined in bot config:
- Uses `config.notify_targets()` from bot configuration
- Supports multiple chat targets (personal + group channels)
- Preserves message thread IDs for grouped chats

### Error Handling

- Invalid JSON: Logged, no reminders loaded
- Invalid reminder config: Skipped, others still loaded
- Failed job execution: Logged, bot continues running
- Scheduler not initialized: Commands show error message

### Validation

Each reminder config is validated for:
- Required fields (id, name, interval_type, message)
- Valid interval_type ("interval" or "cron")
- Positive interval_minutes (for interval type)
- Valid hour (0-23) and minute (0-59)
- At least one of day_of_week or day_of_month (for cron type)

---

## Testing

Run the test suite:
```bash
python test_reminders_system.py
```

Tests covered:
- ✓ Config loading and parsing
- ✓ Reminder validation
- ✓ APScheduler trigger creation
- ✓ Summary generation and display

---

## Troubleshooting

### Reminders not sending
1. Check `/reminderslist` - are they registered?
2. Check bot logs for errors
3. Verify notification targets in bot config
4. Ensure bot has send_message permission

### "Reminders scheduler belum diinisialisasi" error
- Bot is starting up, wait a few seconds
- Or scheduler failed to initialize, check logs

### Invalid JSON error
- Validate reminders_config.json syntax (use online JSON validator)
- Check for missing commas, quotes, brackets

### Reminders not reloading
- Verify reminders_config.json is valid JSON
- Check bot logs for reload errors
- Ensure file is saved before sending `/remindersreload`

---

## Backward Compatibility

The old reminder configuration (REMINDER_HOUR, REMINDER_MINUTE, REMINDER_INTERVAL_MINUTES) is now deprecated but not removed. The system prefers `reminders_config.json`:

```python
# Old config (still supported for backward compat, but not used)
REMINDER_HOUR=17
REMINDER_MINUTE=30
REMINDER_INTERVAL_MINUTES=0

# New config (preferred)
reminders_config.json with dynamic reminders
```

---

## Future Enhancements

Potential improvements:
- [ ] Per-reminder notification target routing
- [ ] Reminder execution history logging
- [ ] Conditional reminders (if X, then send Y)
- [ ] User-specific reminders
- [ ] Web UI for reminder management
- [ ] Reminder persistence across bot restarts

---

## Files

**Core Files:**
- `reminders_config.json` - Configuration file
- `reminders_manager.py` - Config loading & validation
- `reminders_scheduler.py` - APScheduler integration

**Integration:**
- `bot.py` - Bot command handlers & initialization

**Testing:**
- `test_reminders_system.py` - Test suite

---

## Support

For issues or questions:
1. Check this documentation
2. Review logs: `logger.info()` statements
3. Run test suite: `python test_reminders_system.py`
4. Check reminders_config.json validity

---

**Version:** 1.0  
**Last Updated:** August 26, 2026  
**Status:** ✓ Production Ready

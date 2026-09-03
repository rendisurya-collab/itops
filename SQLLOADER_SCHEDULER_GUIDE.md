# SQLLoader & DynamicScheduler - Implementation Guide

## Overview

SQLLoader & DynamicScheduler adalah sistem otomatis untuk memuat file `.sql` dari folder eksternal dan mengeksekusinya berdasarkan jadwal dinamis yang didefinisikan di file konfigurasi JSON.

**Key Features:**
- ✓ Auto-scan folder D:\mybot\tools\queries untuk .sql files
- ✓ Konfigurasi dinamis via query_config.json
- ✓ APScheduler untuk job scheduling (Cron + Interval triggers)
- ✓ Dual output strategy: Text (<=3500 chars) atau Excel export
- ✓ Conditional notification: Hanya kirim Telegram jika query return > 0 rows
- ✓ Hot-reload: Update SQL/config tanpa restart bot

---

## Architecture

### Module Structure

```
bot.py (main)
├── sql_loader.py
│   └── SQLLoader class
│       ├── load_sql_files()
│       ├── load_config()
│       ├── reload_sql_files()
│       ├── validate_query_config()
│       └── list_enabled_queries()
├── query_executor.py
│   └── QueryExecutor class
│       ├── execute_select()
│       ├── format_rows_as_text()
│       ├── export_rows_to_excel()
│       └── process_query_result()
└── dynamic_scheduler.py
    └── DynamicScheduler class
        ├── initialize_scheduler()
        ├── register_jobs_from_config()
        ├── _execute_query_job()
        ├── start_scheduler()
        └── reload_and_reregister()
```

---

## Setup & Configuration

### 1. Environment Variable

File: `.env`

```env
SQL_FOLDER_PATH=D:\mybot\tools\queries
```

### 2. Folder Structure

```
D:\mybot\tools\queries\
├── query_config.json           (config file - required)
├── daily_validation.sql        (query file 1)
├── sync_stock_hourly.sql       (query file 2)
└── [other .sql files...]
```

### 3. Configuration File Format

File: `D:\mybot\tools\queries\query_config.json`

```json
{
  "daily_validation": {
    "enabled": true,
    "schedule_type": "cron",
    "hour": 15,
    "minute": 0,
    "telegram_chat_id": "-1001606974713:33096",
    "description": "Daily validation at 15:00 WIB"
  },
  "sync_stock_hourly": {
    "enabled": true,
    "schedule_type": "interval",
    "interval_minutes": 60,
    "telegram_chat_id": "-1001606974713:33096",
    "description": "Sync stock every hour"
  }
}
```

**Field Descriptions:**

| Field | Type | Required | Description |
|-------|------|----------|-------------|
| `enabled` | bool | Yes | Enable/disable job |
| `schedule_type` | str | Yes | "cron" atau "interval" |
| `hour` | int | If cron | Hour (0-23) untuk cron trigger |
| `minute` | int | If cron | Minute (0-59) untuk cron trigger |
| `interval_minutes` | int | If interval | Interval dalam menit untuk interval trigger |
| `telegram_chat_id` | str | No | Target chat ID (optional, use notify_targets if empty) |
| `description` | str | No | Job description |

---

## SQL File Guidelines

### File Naming

- Use snake_case untuk nama file
- Harus berekstensi `.sql`
- Nama file (tanpa .sql) harus match key di query_config.json

### Query Requirements

- **MUST** berupa SELECT statement
- **MUST** return hasil dalam format tabular
- Jika query return 0 rows → **TIDAK** kirim Telegram notification
- Jika query return > 0 rows → kirim ke Telegram (text atau Excel)

### Example Query

```sql
-- daily_validation.sql
SELECT 
    article_code,
    site_code,
    company_code,
    stock,
    CURRENT_TIMESTAMP AS validation_time
FROM your_table
WHERE status = 'active'
ORDER BY article_code
```

---

## Execution Workflow

### Step 1: Bot Startup

```
main() → app.job_queue.run_once(_init_sql_loader_scheduler)
    ↓
_init_sql_loader_scheduler()
    ├─ SQLLoader.load_sql_files()    [load .sql dari disk]
    ├─ SQLLoader.load_config()       [load query_config.json]
    ├─ DynamicScheduler.initialize_scheduler()
    ├─ DynamicScheduler.start_scheduler()
    └─ DynamicScheduler.register_jobs_from_config()
```

### Step 2: Job Execution (at scheduled time)

```
APScheduler trigger (cron/interval)
    ↓
DynamicScheduler._execute_query_job(query_name)
    ├─ Step A: QueryExecutor.execute_select(query)
    │   └─ if error → log, stop
    │
    ├─ Step B: Check rows
    │   ├─ if 0 rows → log "no data", STOP (no notification)
    │   └─ if > 0 rows → proceed
    │
    ├─ Step C: Process result
    │   └─ QueryExecutor.process_query_result()
    │       ├─ if text <= 3500 chars → return text only
    │       └─ if text > 3500 chars → return text + Excel buffer
    │
    └─ Step D: Send Telegram notification
        ├─ if excel_bytes → send as file attachment
        └─ if no excel → send as text message
```

---

## Usage Examples

### Example 1: Add New Query (Daily Cron)

1. Create SQL file: `D:\mybot\tools\queries\inventory_check.sql`

```sql
SELECT 
    product_id,
    warehouse,
    quantity,
    reorder_level
FROM inventory
WHERE quantity < reorder_level
```

2. Update `query_config.json`:

```json
{
  "inventory_check": {
    "enabled": true,
    "schedule_type": "cron",
    "hour": 9,
    "minute": 30,
    "telegram_chat_id": "-1001606974713",
    "description": "Daily inventory check at 09:30"
  }
}
```

3. Reload on running bot (or restart):

```python
# Via bot command (future feature)
# /queryreload → calls dynamic_scheduler.reload_and_reregister()
```

### Example 2: Add New Query (Hourly Interval)

1. Create SQL file: `D:\mybot\tools\queries\stock_sync.sql`

```sql
SELECT sku, quantity, last_sync FROM stock_log WHERE updated_at > NOW() - INTERVAL 1 HOUR
```

2. Update config:

```json
{
  "stock_sync": {
    "enabled": true,
    "schedule_type": "interval",
    "interval_minutes": 60,
    "telegram_chat_id": "-1001606974713:12345"
  }
}
```

### Example 3: Disable Query

```json
{
  "old_query": {
    "enabled": false,
    "schedule_type": "cron",
    "hour": 12,
    "minute": 0
  }
}
```

---

## Testing

Run the test suite:

```bash
python test_sqlloader_scheduler.py
```

**Test Coverage:**

✓ SQLLoader: file loading, config parsing, validation  
✓ QueryExecutor: text formatting, Excel export, large data handling  
✓ DynamicScheduler: initialization, job registration, trigger setup  

---

## Troubleshooting

### Issue: "SQL folder tidak ditemukan"

**Cause:** `SQL_FOLDER_PATH` tidak dikonfigurasi di `.env`  
**Solution:** Set `SQL_FOLDER_PATH=D:\mybot\tools\queries` di `.env`

### Issue: "Tidak ada query yang enabled"

**Cause:** Semua query di config punya `"enabled": false`  
**Solution:** Set `"enabled": true` di query_config.json untuk query yang ingin dijalankan

### Issue: Query tidak dijalankan pada jadwal

**Cause:** 
- Query filename tidak sesuai config key
- Query syntax error (bukan SELECT statement)
- APScheduler belum start

**Solution:**
- Verify filename match (case-sensitive)
- Check query syntax: harus `SELECT ...`
- Check bot logs untuk error messages

### Issue: Notifikasi tidak terkirim

**Cause:**
- Chat ID salah atau invalid
- Query return 0 rows (by design, no notification sent)
- Telegram API rate limit

**Solution:**
- Verify `telegram_chat_id` di config
- Check bot logs untuk error
- Jika query return 0 rows, no notification is expected behavior

---

## Admin Commands (Future)

```
/querylist              - List semua registered queries & jobs
/queryreload            - Reload SQL files & config dari disk
/queryunregister <name> - Unregister specific job
/queryenable <name>     - Enable query
/querydisable <name>    - Disable query
```

---

## Performance Considerations

- **Query Execution Timeout:** 30 detik (dapat dikonfigurasi)
- **Excel Export:** Optimized untuk 5000+ rows
- **Notification Batch:** Individual notifications per query (tidak di-batch)
- **Memory:** Query results cached in memory selama execution

---

## Version History

- **v1.0** (2026-09-03): Initial release
  - SQLLoader: auto-scan, hot-reload
  - DynamicScheduler: Cron + Interval triggers
  - QueryExecutor: dual output (text/Excel)
  - Telegram notifications dengan conditional logic

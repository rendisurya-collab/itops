# Remaining Refactoring Work: Tasks #4.2, #5, #6, #7

## Status
- ✅ Tasks #1-3: Read-only query handlers completed (4 handlers refactored)
- ✅ Task #4.1: /delreservation refactored to direct execution
- ⏳ Task #4.2: /releasevoucher, /updateshift (same pattern as delreservation)
- ⏭️ Task #5: Jira handlers (/log, /edit, /delete, /myjira)
- ⏭️ Task #6: Guidance handlers (/addguide, /editguide, /delguide)
- ⏭️ Task #7: Final cleanup & testing

---

## Task #4.2: Complete Data Modification Handlers

### /releasevoucher (30 min)

**Current state**: Has ConversationHandler with RELVOUCHER_INPUT, RELVOUCHER_CONFIRM states

**Pattern**: Identical to /delreservation (file or text input → execute)

**Steps**:
1. Rename `releasevoucher_start/input/confirm` → `releasevoucher_command`
2. Consolidate into single handler (copy pattern from delreservation_command)
3. Remove confirmation buttons (no more yes/no callback)
4. Update registration: replace ConversationHandler with CommandHandler + MessageHandlers
5. Test: upload CSV, type manual, verify Excel export on large output

**Risk**: Low (same proven pattern as delreservation)

### /updateshift (30 min)

**Current state**: Has ConversationHandler with UPDATESHIFT_UPLOAD state

**Pattern**: File upload → parse sheet → upsert to Google Sheets

**Steps**:
1. Rename `updateshift_start/upload` → `updateshift_command`
2. Remove confirmation step (currently shows preview + buttons)
3. Update registration: replace ConversationHandler with CommandHandler + MessageHandler
4. Execute immediately on file upload
5. Support text args: `/updateshift 2026-08-26 07:00 15:00 SHIFT_1 Bagus, Tri`

**Risk**: Low (mostly self-contained, Google Sheets integration already tested)

---

## Task #5: Jira Handlers (Complex)

### Overview
- `/log` (5 steps): collect issue → time → description → date → confirm → execute
- `/edit` (3 steps): pick issue → pick worklog → edit fields → execute
- `/delete` (3 steps): pick issue → pick worklog → confirm → execute
- `/myjira` (3 steps): email → token → confirm → save

**Strategy**: Keep multi-step interactive flow, but **remove explicit confirmation button**. Execute on final input.

### /log Refactoring (60 min)

**Current state**: 5 states (LOG_ISSUE, LOG_TIME, LOG_DESC, LOG_DATE, LOG_CONFIRM)

**New flow**:
```
User: /log PROJ-123 2h Work on feature
Bot: executes immediately (shortcut)

OR

User: /log
Bot: Issue key? → user replies
Bot: Duration? → user replies
Bot: Description? → user replies
Bot: Date offset (0=today)? → user replies
Bot: executes immediately (NO confirmation button)
```

**Key changes**:
- Remove `log_confirm` callback handler
- Remove `log_confirm_yes/no` InlineKeyboardButton
- On final field input → execute directly
- Execution logic already exists in `jira_client.add_worklog()`

**Implementation**:
1. Keep multi-state flow in ConversationHandler OR convert to manual message handling
2. Remove yes/no confirmation step
3. Register single CommandHandler for `/log` (not multi-handler system)

**Risk**: Medium (complex state machine, but logic is solid)

### /edit & /delete (similar approach)

Both use worklog selection (inline buttons to pick which worklog to edit/delete). These aren't confirmation buttons — they're selection buttons. **Keep these**, just remove the final yes/no confirmation.

### /myjira (simpler)

Remove confirmation buttons after token validation. Directly save account.

---

## Task #6: Guidance Handlers (Very Complex)

### Overview
- `/addguide`: 10+ states (TITLE → KEYWORDS → CONTENT → ACTION_* → CONFIRM)
- `/editguide`: Menu-driven (implicit flow, hard to track states)
- `/delguide`: Simple pick → confirm → delete

### /delguide (10 min - easiest)

Just remove the confirmation buttons. Change from:
```
Pick guide → confirm [Yes/No]? → delete
```
To:
```
Pick guide → delete immediately
```

### /addguide (90 min - complex)

Keep the field collection (title, keywords, content, action script, etc.), but:
- Remove the final "Save guide?" confirmation button
- On user clicking "Done" or final input → save immediately

**Implementation**:
- Keep ConversationHandler states (too complex to refactor to direct execution)
- Just remove the final confirmation step in `ADD_GUIDE_CONFIRM` state

### /editguide (90 min - most complex)

Menu-driven with implicit save. Currently:
```
User picks guidance → menu shows [Edit Title] [Edit Keywords] ... [Save]
User clicks [Save] → save to store
```

Keep this as-is, remove explicit yes/no buttons if they exist.

---

## Task #7: Final Cleanup & Testing

### Cleanup (30 min)

1. **Identify all orphaned functions**:
   ```bash
   grep "^async def.*_start\|^async def.*_input\|^async def.*_confirm" bot.py
   ```
   Remove any that are no longer called.

2. **Identify unused state constants**:
   - After Tasks #4-6, constants like DELRES_INPUT, LOG_ISSUE, etc. will be orphaned
   - Update the state range assignment (currently `range(41)`)

3. **Verify ConversationHandler count**:
   ```bash
   grep -c "ConversationHandler(" bot.py
   ```
   Should be close to 0-2 after refactoring (maybe just guidance handlers if kept).

4. **Check for orphaned CallbackQueryHandler registrations**:
   - Remove `app.add_handler(CallbackQueryHandler(..., pattern="^delres_"))`
   - Remove similar for other handlers

### Testing (60 min)

1. **Shortcut commands** (if supported):
   - `/stock 8000044321 SS20` → immediate execution ✓
   - `/log PROJ-123 2h Work` → immediate execution ✓

2. **Interactive commands**:
   - `/stock` → ask → user input → execute ✓
   - `/log` → ask fields → user inputs → execute ✓
   - `/delreservation` → ask → user file/text → execute ✓

3. **Error handling**:
   - Invalid args → usage hint shown ✓
   - API errors → try-except message shown ✓
   - File format errors → clear error message ✓

4. **Output formatting**:
   - Small output (<3500 chars) → text message ✓
   - Large output (>3500 chars) → Excel file ✓

5. **End-to-end flows**:
   - `/stock 8000044321 SS20` → bot responds in <5s ✓
   - `/log PROJ-123 2h` → worklog added, confirmation visible ✓
   - `/delreservation` + file → processes all rows, shows result ✓

### Verification checklist

```bash
# 1. Syntax check
python -m py_compile bot.py

# 2. Count handlers
grep -n "ConversationHandler" bot.py | wc -l
# Expected: 0-2 (only complex handlers if not fully refactored)

# 3. Find unused state constants
grep "^    [A-Z_]*INPUT\|^    [A-Z_]*CONFIRM" bot.py | wc -l
# Expected: <10 (mostly cleaned up)

# 4. Test syntax locally
python -c "import bot; print('✓ Bot imports successfully')"
```

---

## Effort Estimate

| Task | Complexity | Time | Notes |
|------|-----------|------|-------|
| #4.2 (releasevoucher) | Low | 30 min | Proven pattern |
| #4.2 (updateshift) | Low | 30 min | Self-contained |
| #5 (/log) | Medium | 60 min | Complex state machine |
| #5 (/edit, /delete) | Medium | 45 min | Keep button-based selection |
| #5 (/myjira) | Low | 15 min | Simple |
| #6 (/delguide) | Low | 10 min | Just remove buttons |
| #6 (/addguide) | High | 90 min | Lots of states |
| #6 (/editguide) | High | 90 min | Menu-driven, implicit |
| #7 (cleanup) | Low | 30 min | Straightforward |
| #7 (testing) | Medium | 60 min | Manual verification |

**Total remaining: ~6-7 hours** (can be parallelized or broken into sessions)

---

## Quick Start: Next Steps

### Option A: Complete Tasks #4.2 + #7 (1-2 hours, get to "mostly working")
1. Copy /delreservation pattern to /releasevoucher
2. Copy /delreservation pattern to /updateshift
3. Clean up orphaned handlers and state constants
4. Quick manual testing

**Result**: 6 handlers fully refactored, ready for production

### Option B: Complete Tasks #4-6 (6-7 hours, full refactoring)
1. Do everything above
2. Refactor all Jira handlers (keep multi-step, remove confirmation)
3. Refactor all guidance handlers (mostly keep as-is, remove buttons)
4. Full testing suite

**Result**: All handlers refactored, production-ready

### Option C: Minimal viable (30 min)
Just finish #4.2 (/releasevoucher, /updateshift), push to main. Leave #5-6 as "future work".

---

## Code Examples for Reference

### Pattern: Direct execution with file/text input
```python
@restricted
async def releasevoucher_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    chat_id = update.effective_chat.id
    thread_id = _thread_id_from_update(update)
    
    if update.message.document:
        # File upload → process immediately
        doc = update.message.document
        file = await doc.get_file()
        csv_path = await download_and_save(file)
        output = await asyncio.to_thread(_execute_releasevoucher, csv_path)
        await send_result(context.bot, chat_id, output, thread_id)
    elif update.message.text:
        # Text input → parse, process immediately
        ids = parse_voucher_ids(update.message.text)
        output = await asyncio.to_thread(_execute_releasevoucher_ids, ids)
        await send_result(context.bot, chat_id, output, thread_id)
    else:
        # Initial /releasevoucher → ask for input
        await context.bot.send_message(chat_id, "Upload CSV or type IDs...", ...)
```

### Pattern: Multi-step without confirmation (Jira)
```python
@restricted
async def log_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    # Parse shortcut if provided
    if context.args:
        issue, time, desc, date = parse_shortcut(context.args)
        result = await jira_client.add_worklog(issue, time, desc, date)
        await send_result(context.bot, chat_id, result)
        return
    
    # Interactive multi-step (keep state machine, just remove confirmation)
    await update.message.reply_text("Issue key?")
    # → User replies → ask next field → ... → on final field → execute
```

---

## Related Commits

- `be82d6c`: Tasks #1-3 (read-only queries)
- `4d9c240`: Task #4.1 (/delreservation)
- `[next]`: Task #4.2-7 (remaining work)


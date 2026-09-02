# Direct Execution Handler Refactoring Template

## Overview
Convert all command handlers from **ConversationHandler state machine + confirmation prompts** to **Direct Execution Pattern** (instant parsing → execute → result).

## Pattern Comparison

### OLD PATTERN (State Machine)
```
/stock 8000044321 SS20
    ↓ (checkstock_start)
[State: CHECKSTOCK_INPUT] Ask user for input
    ↓ (checkstock_input)
[State: CHECKSTOCK_CONFIRM] Show preview + [Yes] [No] buttons
    ↓ (checkstock_confirm callback)
Execute & Send result
```

### NEW PATTERN (Direct Execution)
```
/stock 8000044321 SS20
    ↓ (stock_command)
Parse args/text
    ↓
Validate
    ↓ 
Execute directly (no prompt)
    ↓
Send result (text or Excel)
```

## Key Differences

| Aspect | Old | New |
|--------|-----|-----|
| **Handlers** | 3 functions (start, input, confirm) | 1 function |
| **States** | Multiple state constants | None |
| **Confirmation** | InlineKeyboardMarkup buttons | None |
| **User steps** | 2-3 steps | 1 step |
| **Registration** | ConversationHandler | CommandHandler |
| **Context.user_data** | Store state data | Not used (inline parsing) |

## Implementation Guide

### Step 1: Identify the Handler
- Handler command: `/stock`
- Current functions: `checkstock_start`, `checkstock_input`, `checkstock_confirm`
- State constants: `CHECKSTOCK_INPUT`, `CHECKSTOCK_CONFIRM`

### Step 2: Consolidate Logic
Extract common parsing & execution from the 3 functions into a single handler:

```python
@restricted
async def stock_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Direct execution: /stock [sku] [source] [unit]"""
    
    # Get command text
    text = update.message.text.strip()
    
    # Try parse shortcut args first
    parts = text.split(maxsplit=1)
    if len(parts) > 1 and parts[1].strip():
        args = parts[1].strip().split()
        if len(args) >= 2:
            data = {
                "article_ids": [args[0]],
                "source_ids": [args[1]],
                "units": [args[2]] if len(args) > 2 else ["EA"],
            }
        else:
            # Not enough args
            await update.message.reply_text(
                "❌ Format salah. Gunakan: /stock [SKU] [SOURCE] [UNIT]"
            )
            return
    else:
        # No args, ask for input once (interactive, no state machine)
        await update.message.reply_text(
            "<b>📦 Check Stock (OAA)</b>\n\n"
            "Kirim format:\n"
            "<code>sku: 8000044321\n"
            "source: SS20</code>\n\n"
            "Atau shortcut: <code>/stock 8000044321 SS20</code>",
            parse_mode=ParseMode.HTML,
        )
        return
    
    # Execute directly (no confirmation step)
    try:
        result = await _execute_stock_check(data)
        await update.message.reply_text(result, parse_mode=ParseMode.HTML)
    except Exception as e:
        await update.message.reply_text(f"❌ Error: {html.escape(str(e))}")
```

### Step 3: Extract Helper Functions
- Keep existing `_parse_checkstock_text()` ✓
- Keep existing `_format_stock_response()` ✓
- Create `_execute_stock_check()` from old `checkstock_confirm` logic

### Step 4: Remove Old Code
- Delete: `checkstock_start`, `checkstock_input`, `checkstock_confirm`
- Delete: state constants `CHECKSTOCK_INPUT`, `CHECKSTOCK_CONFIRM`
- Remove from `add_handlers()`: old ConversationHandler registration

### Step 5: Register New Handler
In `main()` or `add_handlers()`:
```python
app.add_handler(CommandHandler("stock", stock_command))
```

---

## Handler Categories & Priority

### Category A: Read-Only Queries (Lower Risk)
Minimal logic, instant result, no side effects:
- [ ] `/cekpromo` (/promo)
- [ ] `/cekawb` (/awb)
- [ ] `/awbjne`
- [ ] `/cekpromo_start`, `cekpromo_input`

**Why:** Users can safely cancel/retry, no confirmation needed.

### Category B: Data Modifications (High Priority)
Write operations, require caution but no real confirmation needed (try-except handles):
- [ ] `/stock` (read-only, but high traffic)
- [ ] `/delreservation`
- [ ] `/releasevoucher`
- [ ] `/updateshift`

**Why:** These MUST be fast. Try-except protects against errors.

### Category C: Complex Multi-Step (Complex Refactor)
State-heavy, dynamic fields:
- [ ] `/log` (Jira worklog with date parsing)
- [ ] `/edit` (worklog selection)
- [ ] `/delete` (worklog selection)
- [ ] `/myjira` (token validation, account save)
- [ ] `/addguide`, `/editguide`, `/delguide`
- [ ] `/run` (dynamic param collection)

**Why:** These need careful arg parsing, but can still be direct (support multi-input fallback, no confirm buttons).

---

## Error Handling Strategy

### Old Pattern
```python
# In confirmation callback
try:
    # execute
except Exception as e:
    await send_error()
```

### New Pattern  
```python
# In direct handler
try:
    # parse
    # execute
    # send result
except ValueError as e:
    # Parsing error -> usage hint
    await send_usage_hint()
except Exception as e:
    # Execution error -> technical error message
    await send_error_message()
```

---

## Output Formatting

For both old & new patterns:

```python
# If output <= 3500 chars: send as HTML text
if len(output) <= 3500:
    await bot.send_message(text=formatted_output, parse_mode=ParseMode.HTML)

# If output > 3500 chars: export to Excel in-memory, send as file
else:
    excel_buffer = _build_log_excel_bytes(output)
    await bot.send_document(document=InputFile(excel_buffer, filename=...))
```

---

## Checklist for Each Refactoring

- [ ] Identify all 3 handler functions
- [ ] Copy logic from all 3 into unified function
- [ ] Extract parsing to helper (if not already exists)
- [ ] Extract execution to `_execute_*()` helper
- [ ] Add error handling: ValueError (usage) + Exception (technical)
- [ ] Verify output formatting (text/Excel dual)
- [ ] Test shortcut args parsing
- [ ] Test interactive fallback (multi-line input)
- [ ] Remove old state constants
- [ ] Remove old handlers from ConversationHandler registration
- [ ] Register new CommandHandler
- [ ] Manual test: shortcut mode, interactive mode, error cases
- [ ] Commit with message: `refactor: convert /command to direct execution`

---

## Example Commits (as we complete each)

```
refactor: convert /stock to direct execution
refactor: convert /delreservation to direct execution
refactor: convert /log to direct execution
...
refactor: remove all ConversationHandler state machine registrations
```

---

## Final Validation

After all conversions:
1. Run `python -m py_compile bot.py`
2. Start bot locally, test each command with shortcut + interactive modes
3. Verify error messages are clear (no truncation)
4. Verify large output → Excel export works
5. Check ConversationHandler is not used (grep should return 0 results except for legacy code)



---

## COMPLETED: Tasks #1-3 ✅

### Summary
- **Task #1**: Template utilities built (parse_shortcut_args, execute_safe, @direct_exec_handler, etc.)
- **Task #2**: /promo, /awb, /awbjne refactored (removed ~6 handler functions, 3 state constants)
- **Task #3**: /stock refactored (single handler, Excel export on large output)

**Commit**: `be82d6c` — "refactor: implement direct-execution pattern for read-only queries"

---

## NEXT: Tasks #4-7

### Task #4: Data Modification Handlers (/delreservation, /releasevoucher, /updateshift)

**Current Pattern (State Machine):**
```
/delreservation
  ↓ (delreservation_start)
[State: DELRES_INPUT] → User uploads CSV or types text
  ↓ (delreservation_input)
[State: DELRES_CONFIRM] → Show preview + [✅ Yes] [❌ No] buttons
  ↓ (delreservation_confirm callback)
Execute & Send result
```

**New Pattern (Direct Execution):**
```
/delreservation
  ↓ (delreservation_command)
Ask for input (ONE step, no state)
  ↓ User sends file/text
Execute directly (no confirmation)
  ↓
Send result (text or Excel)
```

**Implementation Strategy:**
1. **Consolidate 3 handlers into 1:**
   - `delreservation_start` + `delreservation_input` + `delreservation_confirm` → `delreservation_command`
   - `releasevoucher_start` + `releasevoucher_input` + `releasevoucher_confirm` → `releasevoucher_command`
   - `updateshift_start` + `updateshift_input` + `updateshift_upload` → `updateshift_command` (already partially direct)

2. **Remove confirmation InlineKeyboardMarkup** (the [✅ Eksekusi] [❌ Batal] buttons)

3. **Execution logic already solid** (has try-except, Excel export). Just move into main handler.

4. **Keep user-facing behavior intuitive:**
   - User still uploads file → bot processes immediately
   - User still types text → bot processes immediately
   - No confirmation step needed (try-except protects)

**Files to modify:**
- Remove DELRES_INPUT, DELRES_CONFIRM, RELVOUCHER_INPUT, RELVOUCHER_CONFIRM state constants
- Update handler registration (remove ConversationHandler, add CommandHandler)
- Handlers: delreservation_*, releasevoucher_*, updateshift_*

**Risk:** Medium (data modification, but protected by try-except and temp file cleanup)

---

### Task #5: Jira Handlers (/log, /edit, /delete, /myjira)

**Current Pattern (Complex State Machines):**
- `/log`: LOG_ISSUE → LOG_TIME → LOG_DESC → LOG_DATE → LOG_CONFIRM (5+ steps)
- `/edit`: PICK_ISSUE_FOR_EDIT → PICK_WORKLOG_EDIT → EDIT_TIME → EDIT_DESC (4 steps)
- `/delete`: PICK_ISSUE_FOR_DELETE → PICK_WORKLOG_DELETE → CONFIRM_DELETE (3 steps)
- `/myjira`: MYJIRA_EMAIL → MYJIRA_TOKEN → MYJIRA_CONFIRM (3 steps)

**New Pattern (Direct Execution):**
- Support shortcut args if provided: `/log PROJ-123 2h Work on feature`
- If args missing: ask interactively ONE field at a time (still no state machine)
- Execute on each input, no final confirmation

**Key Challenge:**
- Jira handlers have multi-field logic (issue key, time, description, date)
- Current approach uses ConversationHandler to collect fields step-by-step
- New approach: still collect fields interactively, but WITHOUT state machine (use message handlers per-context)

**Implementation Strategy:**
1. For `/log` shortcut: `/log PROJ-123 2h` → parse, execute immediately
2. For `/log` interactive: ask "Issue key?" → user replies → ask "Duration?" → user replies → execute
3. Use `context.user_data` to track collection stage, but WITHOUT state constants
4. On each message, check if we have all needed fields → execute

**Alternative (Simpler):**
- Keep the interactive multi-step flow, but **remove the confirmation button**
- After collecting all fields, execute immediately (no yes/no prompt)
- This is still "direct execution" — just with multi-message input gathering

**Risk:** High (complex logic, multiple fields, Jira API integration)

**Recommendation:** Start with "Alternative" approach (simpler) — keep multi-step gathering, just remove confirmation step.

---

### Task #6: Guidance Handlers (/addguide, /editguide, /delguide)

**Current Pattern (Extreme State Machine):**
- `/addguide`: 10+ states (TITLE → KEYWORDS → CONTENT → ACTION_ASK → ACTION_SCRIPT → ACTION_FLAG → ACTION_MODE → ACTION_TYPE → CONFIRM)
- `/editguide`: Menu-driven (pick guide → show menu → edit field → return to menu → save)
- `/delguide`: Simple (pick guide → confirm → delete)

**New Pattern (Direct Execution):**
- For `/delguide`: Only 1 state (pick → execute, no confirm)
- For `/addguide`: Still multi-step interactive, but **no final confirmation button**
- For `/editguide`: Menu-based (still interactive), **no final confirmation**

**Implementation Strategy:**
1. `/delguide` → easy (just remove confirm button)
2. `/addguide` → harder (remove final confirmation, execute when "done" clicked)
3. `/editguide` → hardest (implicit save, no confirm)

**Risk:** Very High (most complex handler, lots of state, guidance store operations)

**Recommendation:** Keep menu-driven approach for `/editguide`, just remove explicit confirmation. For `/addguide`, remove final confirmation step.

---

### Task #7: Final Cleanup & Testing

**Cleanup:**
1. Remove all unused handler functions (old *_start, *_input, *_confirm, *_execute)
2. Remove all unused state constants
3. Verify ConversationHandler count → should be very few (maybe 0-1 for updateshift if not refactored)
4. Check for orphaned ConversationHandler states → clean up

**Testing Strategy:**
1. **Local bot testing:**
   - Test each command with shortcut args (if supported)
   - Test each command with interactive input
   - Verify error handling (try-except messages clear)
   - Verify Excel export works if output > 3500 chars

2. **Verification checklist:**
   ```bash
   # Check for remaining ConversationHandler usages
   grep -n "ConversationHandler" bot.py | grep -v "#"
   
   # Check for remaining state constants
   grep -n "^    [A-Z_]*INPUT\|^    [A-Z_]*CONFIRM" bot.py
   
   # Check syntax
   python -m py_compile bot.py
   ```

3. **Manual test cases:**
   - `/stock 8000044321 SS20` → executes immediately ✓
   - `/stock` → asks for input, user types, executes ✓
   - `/delreservation` → asks for file/text, user uploads, executes ✓
   - `/log PROJ-123 2h` → executes immediately ✓
   - `/log` → asks for issue, time, desc, date, then executes ✓
   - Error cases: `/stock SK` → usage hint shown ✓
   - Large output: check Excel export works ✓

**Final Commit:**
- Combine all refactored handlers into single commit: `refactor: complete direct-execution pattern for all handlers`

---

## Estimated Effort

| Task | Complexity | Time | Priority |
|------|-----------|------|----------|
| #4 (Data mods) | Medium | 1-2 hrs | High |
| #5 (Jira) | High | 2-3 hrs | Medium |
| #6 (Guidance) | Very High | 2-3 hrs | Low |
| #7 (Cleanup) | Low | 30 min | High |

**Recommendation:** Complete #4 (quick win), then #7 (cleanup), then decide on #5 and #6 based on time/priority.


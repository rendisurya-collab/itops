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


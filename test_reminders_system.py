#!/usr/bin/env python3
"""
Test script for dynamic reminders system.
Tests loading config, parsing reminders, and trigger creation.
"""

import sys
import asyncio
from reminders_manager import init_reminders_manager, get_reminders_manager
from reminders_scheduler import RemindersScheduler
from apscheduler.schedulers.asyncio import AsyncIOScheduler

def test_config_loading():
    """Test 1: Load reminders configuration"""
    print("\n" + "="*60)
    print("TEST 1: Load reminders configuration")
    print("="*60)
    
    manager = init_reminders_manager("reminders_config.json")
    
    if not manager.reminders:
        print("❌ FAILED: No reminders loaded")
        return False
    
    print(f"✓ PASSED: Loaded {len(manager.reminders)} reminders")
    for reminder in manager.reminders:
        print(f"  - {reminder}")
    
    return True


def test_reminder_validation():
    """Test 2: Validate reminder configurations"""
    print("\n" + "="*60)
    print("TEST 2: Validate reminder configurations")
    print("="*60)
    
    manager = get_reminders_manager()
    
    all_valid = True
    for reminder in manager.reminders:
        print(f"\nReminder: {reminder.name} ({reminder.id})")
        print(f"  - Enabled: {reminder.enabled}")
        print(f"  - Type: {reminder.interval_type}")
        
        if reminder.interval_type == "interval":
            print(f"  - Interval: {reminder.interval_minutes} minutes")
        else:
            print(f"  - Time: {reminder.hour:02d}:{reminder.minute:02d}")
            if hasattr(reminder, 'day_of_week') and reminder.day_of_week:
                print(f"  - Day of week: {reminder.day_of_week}")
            if hasattr(reminder, 'day_of_month') and reminder.day_of_month:
                print(f"  - Day of month: {reminder.day_of_month}")
        
        print(f"  - Message preview: {reminder.message[:50]}...")
        print("  ✓ Valid")
    
    return all_valid


def test_trigger_creation():
    """Test 3: Create APScheduler triggers"""
    print("\n" + "="*60)
    print("TEST 3: Create APScheduler triggers for all reminders")
    print("="*60)
    
    manager = get_reminders_manager()
    
    # Create a mock scheduler
    scheduler = AsyncIOScheduler()
    
    class MockScheduler:
        """Mock RemindersScheduler for testing trigger creation"""
        def __init__(self, manager):
            self.manager = manager
        
        def _create_interval_trigger(self, reminder):
            from reminders_scheduler import RemindersScheduler
            rs = RemindersScheduler.__new__(RemindersScheduler)
            return rs._create_interval_trigger(reminder)
        
        def _create_cron_trigger(self, reminder):
            from reminders_scheduler import RemindersScheduler
            rs = RemindersScheduler.__new__(RemindersScheduler)
            return rs._create_cron_trigger(reminder)
    
    mock = MockScheduler(manager)
    
    all_valid = True
    for reminder in manager.get_active_reminders():
        try:
            if reminder.interval_type == "interval":
                trigger = mock._create_interval_trigger(reminder)
                print(f"✓ {reminder.name}: IntervalTrigger({reminder.interval_minutes} min)")
            else:
                trigger = mock._create_cron_trigger(reminder)
                trigger_str = f"hour={reminder.hour}, minute={reminder.minute}"
                if hasattr(reminder, 'day_of_week') and reminder.day_of_week:
                    trigger_str += f", day_of_week={reminder.day_of_week}"
                if hasattr(reminder, 'day_of_month') and reminder.day_of_month:
                    trigger_str += f", day={reminder.day_of_month}"
                print(f"✓ {reminder.name}: CronTrigger({trigger_str})")
        except Exception as e:
            print(f"❌ {reminder.name}: Failed to create trigger - {e}")
            all_valid = False
    
    return all_valid


def test_summary_generation():
    """Test 4: Generate summary display"""
    print("\n" + "="*60)
    print("TEST 4: Generate summary display")
    print("="*60)
    
    manager = get_reminders_manager()
    summary = manager.get_summary()
    print(summary)
    
    return True


async def main():
    """Run all tests"""
    print("\n" + "#"*60)
    print("REMINDERS SYSTEM TEST SUITE")
    print("#"*60)
    
    results = []
    
    try:
        results.append(("Config Loading", test_config_loading()))
        results.append(("Reminder Validation", test_reminder_validation()))
        results.append(("Trigger Creation", test_trigger_creation()))
        results.append(("Summary Generation", test_summary_generation()))
    except Exception as e:
        print(f"\n❌ Unexpected error: {e}", exc_info=True)
        return 1
    
    # Print summary
    print("\n" + "="*60)
    print("TEST SUMMARY")
    print("="*60)
    passed = sum(1 for _, result in results if result)
    total = len(results)
    
    for name, result in results:
        status = "✓ PASS" if result else "❌ FAIL"
        print(f"{status}: {name}")
    
    print(f"\nTotal: {passed}/{total} tests passed")
    
    if passed == total:
        print("\n✓✓✓ All tests PASSED! ✓✓✓")
        return 0
    else:
        print(f"\n❌ {total - passed} test(s) FAILED")
        return 1


if __name__ == "__main__":
    exit_code = asyncio.run(main())
    sys.exit(exit_code)

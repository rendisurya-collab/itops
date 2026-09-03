"""
Test script untuk verify SQLLoader & DynamicScheduler modules.
"""
import sys
import logging
from pathlib import Path
from sql_loader import SQLLoader
from query_executor import QueryExecutor

# Setup logging
logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO
)
logger = logging.getLogger(__name__)


def test_sql_loader():
    """Test SQLLoader functionality."""
    print("\n" + "="*60)
    print("TEST 1: SQLLoader - Load SQL files from disk")
    print("="*60)

    sql_folder = Path(r"D:\mybot\tools\queries")
    loader = SQLLoader(str(sql_folder))

    # Test load_sql_files
    print(f"\n[TEST] Loading SQL files from: {sql_folder}")
    count = loader.load_sql_files()
    print(f"✓ Loaded {count} SQL files")
    assert count > 0, "No SQL files loaded!"

    # Test get_all_queries
    all_queries = loader.get_all_queries()
    print(f"✓ Retrieved {len(all_queries)} queries from memory")
    for name, content in all_queries.items():
        print(f"  - {name}: {len(content)} chars")

    # Test load_config
    print(f"\n[TEST] Loading config from: {sql_folder / 'query_config.json'}")
    config_loaded = loader.load_config()
    print(f"✓ Config loaded: {config_loaded}")

    config = loader.get_config()
    print(f"✓ Retrieved config with {len(config)} entries")
    for name, entry in config.items():
        print(f"  - {name}: enabled={entry.get('enabled')}, type={entry.get('schedule_type')}")

    # Test list_enabled_queries
    enabled = loader.list_enabled_queries()
    print(f"\n✓ Enabled queries: {enabled}")

    # Test validate_query_config
    for query_name in enabled:
        is_valid, error = loader.validate_query_config(query_name)
        print(f"  - {query_name}: valid={is_valid}, error={error if error else 'OK'}")

    print("\n✓ SQLLoader tests PASSED")
    return loader


def test_query_executor(loader):
    """Test QueryExecutor functionality."""
    print("\n" + "="*60)
    print("TEST 2: QueryExecutor - Execute & format queries")
    print("="*60)

    executor = QueryExecutor(db_connection=None)

    # Get a sample query
    queries = loader.get_all_queries()
    if not queries:
        print("⚠ No queries available, skipping executor tests")
        return executor

    test_query_name = list(queries.keys())[0]
    test_query = queries[test_query_name]

    print(f"\n[TEST] Validating query format: {test_query_name}")
    print(f"Query preview: {test_query[:100]}...")

    # Validate SELECT check
    success, error = test_query.strip().lower().startswith("select"), None
    if not success:
        print(f"⚠ Query is not a SELECT statement, skipping execute test")
        print(f"✓ QueryExecutor format validation works")
        return executor

    print(f"✓ Query format is valid (SELECT)")

    # Test format_rows_as_text with sample data
    print(f"\n[TEST] Testing text formatting with sample data")
    sample_rows = [
        ("ABC123", "SS20", "COMPANY1", 100),
        ("ABC456", "SS21", "COMPANY1", 250),
    ]
    text, exceeded = executor.format_rows_as_text(sample_rows, max_chars=3500)
    print(f"✓ Text formatted: {len(text)} chars, exceeded={exceeded}")
    print(f"Sample output:\n{text[:200]}...")

    # Test format with large dataset
    print(f"\n[TEST] Testing text formatting with large dataset")
    large_rows = [
        (f"SKU_{i:05d}", f"WH_{i % 5}", f"COMPANY_{i % 3}", i * 100)
        for i in range(1, 100)
    ]
    text_large, exceeded_large = executor.format_rows_as_text(large_rows, max_chars=500)
    print(f"✓ Large dataset formatted: {len(text_large)} chars, exceeded={exceeded_large}")
    assert exceeded_large, "Should exceed 500 char limit with 100 rows"
    print(f"✓ Limit exceeded check works correctly")

    # Test process_query_result with small data (no Excel)
    print(f"\n[TEST] Testing process_query_result (small data)")
    text_msg, excel_bytes, error = executor.process_query_result(
        sample_rows,
        "test_query",
        max_text_chars=3500
    )
    print(f"✓ Result processed: text={len(text_msg)} chars, excel={excel_bytes is not None}, error={error}")
    assert excel_bytes is None, "Should not generate Excel for small data"

    # Test process_query_result with large data (with Excel)
    print(f"\n[TEST] Testing process_query_result (large data)")
    text_msg_large, excel_bytes_large, error_large = executor.process_query_result(
        large_rows,
        "test_query_large",
        max_text_chars=500
    )
    print(f"✓ Result processed: text={len(text_msg_large)} chars, excel={excel_bytes_large is not None}, error={error_large}")
    assert excel_bytes_large is not None, "Should generate Excel for large data"
    print(f"✓ Excel generated: {len(excel_bytes_large)} bytes")

    # Test empty result
    print(f"\n[TEST] Testing with empty result")
    text_empty, excel_empty, error_empty = executor.process_query_result(
        [],
        "empty_query",
        max_text_chars=3500
    )
    print(f"✓ Empty result handled: text='{text_empty}'")
    assert "Tidak ada data" in text_empty, "Should handle empty result"

    print("\n✓ QueryExecutor tests PASSED")
    return executor


def test_dynamic_scheduler():
    """Test DynamicScheduler initialization (without running)."""
    print("\n" + "="*60)
    print("TEST 3: DynamicScheduler - Initialization")
    print("="*60)

    from dynamic_scheduler import DynamicScheduler

    sql_folder = Path(r"D:\mybot\tools\queries")
    loader = SQLLoader(str(sql_folder))
    loader.load_sql_files()
    loader.load_config()

    executor = QueryExecutor(db_connection=None)

    print(f"\n[TEST] Creating DynamicScheduler instance")
    scheduler = DynamicScheduler(
        sql_loader=loader,
        query_executor=executor,
        telegram_notify_callback=None  # No Telegram during test
    )
    print(f"✓ DynamicScheduler created")

    print(f"\n[TEST] Initializing AsyncIOScheduler")
    result = scheduler.initialize_scheduler()
    print(f"✓ Scheduler initialized: {result}")
    assert result, "Failed to initialize scheduler"

    print(f"\n[TEST] Registering jobs from config (no start)")
    registered = scheduler.register_jobs_from_config()
    print(f"✓ Registered {registered} jobs")

    print(f"\n[TEST] Checking registered jobs")
    jobs = scheduler.get_registered_jobs()
    print(f"✓ Retrieved {len(jobs)} jobs:")
    for job in jobs:
        print(f"  - {job['name']}: {job['trigger']}")

    print("\n✓ DynamicScheduler tests PASSED")


def main():
    """Run all tests."""
    print("\n" + "#"*60)
    print("# SQLLoader & DynamicScheduler - Integration Test Suite")
    print("#"*60)

    try:
        # Test 1: SQLLoader
        loader = test_sql_loader()

        # Test 2: QueryExecutor
        executor = test_query_executor(loader)

        # Test 3: DynamicScheduler
        test_dynamic_scheduler()

        print("\n" + "#"*60)
        print("# ALL TESTS PASSED ✓")
        print("#"*60 + "\n")
        return 0

    except AssertionError as e:
        print(f"\n✗ TEST FAILED: {e}")
        return 1
    except Exception as e:
        print(f"\n✗ ERROR: {e}")
        import traceback
        traceback.print_exc()
        return 1


if __name__ == "__main__":
    sys.exit(main())

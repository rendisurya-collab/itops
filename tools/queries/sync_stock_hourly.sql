-- Hourly stock sync query (sample)
-- This is a placeholder SELECT query for testing
SELECT 
    'SKU_A001' AS sku,
    'Warehouse_A' AS warehouse,
    500 AS total_stock,
    CURRENT_TIMESTAMP AS last_updated
UNION ALL
SELECT 
    'SKU_B002' AS sku,
    'Warehouse_B' AS warehouse,
    1200 AS total_stock,
    CURRENT_TIMESTAMP AS last_updated
UNION ALL
SELECT 
    'SKU_C003' AS sku,
    'Warehouse_C' AS warehouse,
    750 AS total_stock,
    CURRENT_TIMESTAMP AS last_updated

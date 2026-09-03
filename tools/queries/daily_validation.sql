-- Daily validation query (sample)
-- This is a placeholder SELECT query for testing
SELECT 
    'test_article_001' AS article_code,
    'SS001' AS site_code,
    'COMPANY1' AS company_code,
    100 AS stock,
    CURRENT_TIMESTAMP AS validation_time
UNION ALL
SELECT 
    'test_article_002' AS article_code,
    'SS002' AS site_code,
    'COMPANY1' AS company_code,
    250 AS stock,
    CURRENT_TIMESTAMP AS validation_time

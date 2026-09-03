-- Test query yang return 0 rows (should not send notification)
SELECT 
    'article' AS article_code,
    'site' AS site_code,
    'company' AS company_code,
    0 AS stock
WHERE 1=0

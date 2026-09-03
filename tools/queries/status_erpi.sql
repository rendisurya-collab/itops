SELECT filters, transaction_title, status_code, response, duration, created_at, id
FROM log_api_orders
WHERE status_code != '200' 
  AND created_at >= CURDATE() - INTERVAL 5 DAY
ORDER BY id ASC
LIMIT 50;

SELECT 
    filters,
    COUNT(*) AS sum
FROM log_api_orders 
WHERE transaction_title = 'OMS|Push Sales Order'  AND created_at >= CURDATE() - INTERVAL 2 DAY
GROUP BY filters 
HAVING COUNT(*) >= 2
ORDER BY sum DESC;

SELECT 
    order_ref_no,
    COUNT(*) AS sum
FROM so_status 
WHERE created_at >= CURDATE() - INTERVAL 2 DAY
 AND so_status_code !='C1'
GROUP BY order_ref_no 
HAVING COUNT(*) >= 2
ORDER BY sum DESC;
SELECT `filters`, `transaction_title`, `status_code`, `response`, `duration`, `created_at`, `id`
FROM `log_api_orders`
WHERE `status_code` != '200' 
  AND `created_at` >= CURDATE() - INTERVAL 2 DAY
ORDER BY `id` ASC
LIMIT 50;
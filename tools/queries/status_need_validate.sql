SELECT a.`order_number`, a.`order_date`, a.`order_type`, a.`site_code`, remark_notes, b.updated_at
FROM `mp_data` a
LEFT JOIN mp_remark b on a.order_number = b.order_number
WHERE `validation_status` = '2'
ORDER BY a.`order_date` ASC
LIMIT 10
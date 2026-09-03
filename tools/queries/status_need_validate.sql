SELECT a.`order_number`, a.`order_date`, a.`order_type`, a.`site_code`, remark_notes, b.updated_at
FROM `mp_data` a
LEFT JOIN mp_remark b on a.order_number = b.order_number
WHERE `validation_status` = '2'
ORDER BY a.`order_date` ASC
LIMIT 10;

UPDATE mp_data
SET validation_status = '1',
    validation_is_need_validated = '0'
WHERE order_number IN (
    SELECT order_number FROM (
        SELECT order_number
        FROM mp_data
        WHERE validation_status = '2' 
          AND validation_is_need_validated = '1'
        ORDER BY order_date ASC
        LIMIT 1
    ) AS tmp
);
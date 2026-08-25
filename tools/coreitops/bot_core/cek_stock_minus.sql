SELECT `sku`, `name`,`category`, `brand`, mp_name, `site_code_stock` as sitecode_stock, qty as stock_mapping, `qty_store` as stock_im3
FROM `master_product` a
LEFT JOIN product_sku_mapping b on a.sku = b.mp_sku
WHERE a.wms_source = b.wms_source
AND b.qty > a.qty_store 
AND mp_name NOT LIKE 'RETUR%' AND mp_name NOT LIKE 'KEEP%'
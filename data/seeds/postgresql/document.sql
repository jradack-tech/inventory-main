set client_encoding to 'UTF-8';
INSERT INTO masterdata_document(id, name, term, taxation) VALUES
(1, '出庫手数料（売上）', '売上高', '課税売上10%'),
(2, '元払発送料（売上）', '売上高', '課税売上10%'),
(3, '出庫手数料（仕入）', '仕入高', '課対仕入10%'),
(4, '元払発送料（仕入）', '仕入高', '課対仕入10%'),
(5, 'あんしん決済', '仕入高', '非課税');

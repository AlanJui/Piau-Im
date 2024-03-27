CREATE VIEW han_ji_view_tc AS 
SELECT id, han_ji AS "漢字", chu_im AS "注音", freq AS "常用度" 
FROM han_ji_dict;
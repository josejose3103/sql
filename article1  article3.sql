SELECT
    t1.id AS t1_id,
    t1.day AS t1_day,
    t1.noukamei AS t1_noukamei,
    t1.byoumei AS t1_byoumei,
    t1.kingaku AS t1_kingaku,
    t2.id AS t2_id,
    t2.day AS t2_day,
    t2.noukamei AS t2_noukamei,
    t2.byoumei AS t2_byoumei,
    t2.kingaku AS t2_kingaku
FROM
    article1 AS t1
LEFT JOIN
    article3 AS t2
ON
    t1.noukamei = t2.noukamei
    AND t1.day = t2.day
WHERE
    t1.noukamei = '坂上孝行'
    AND t1.day BETWEEN '2023-03-26' AND '2023-04-10'

UNION

-- article3 を基準にして article1 を結合（LEFT JOIN）
SELECT
    t1.id AS t1_id,
    t1.day AS t1_day,
    t1.noukamei AS t1_noukamei,
    t1.byoumei AS t1_byoumei,
    t1.kingaku AS t1_kingaku,
    t2.id AS t2_id,
    t2.day AS t2_day,
    t2.noukamei AS t2_noukamei,
    t2.byoumei AS t2_byoumei,
    t2.kingaku AS t2_kingaku
FROM
    article3 AS t2
LEFT JOIN
    article1 AS t1
ON
    t1.noukamei = t2.noukamei
    AND t1.day = t2.day
WHERE
    t2.noukamei = '坂上孝行'
    AND (t1.id IS NULL)
    AND t2.day BETWEEN '2023-03-26' AND '2023-04-10';
    SELECT
    '坂上孝行' AS noukamei,
    SUM(IFNULL(t1.kingaku, 0)) + SUM(IFNULL(t2.kingaku, 0)) AS total_kingaku
FROM
    article1 AS t1
LEFT JOIN
    article3 AS t2
ON
    t1.noukamei = t2.noukamei
    AND t1.day = t2.day
WHERE
    t1.noukamei = '坂上孝行'
    AND t1.day BETWEEN '2023-03-26' AND '2023-04-10'

UNION ALL

SELECT
    '坂上孝行' AS noukamei,
    SUM(IFNULL(t2.kingaku, 0)) AS total_kingaku
FROM
    article3 AS t2
LEFT JOIN
    article1 AS t1
ON
    t1.noukamei = t2.noukamei
    AND t1.day = t2.day
WHERE
    t2.noukamei = '坂上孝行'
    AND t1.id IS NULL
    AND t2.day BETWEEN '2023-03-26' AND '2023-04-10';

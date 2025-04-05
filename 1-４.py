SELECT 
    EXTRACT(YEAR FROM day) AS `年`,
    SUM(kingaku) AS `合計`
FROM article3 
WHERE MONTH(day) = 4
  AND DAY(day) BETWEEN 1 AND 4
GROUP BY `年`
ORDER BY `年`;
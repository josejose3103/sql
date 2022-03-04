LOAD DATA LOCAL INFILE 'C:/Users/tsuboi/OneDrive/Documents/table1.cvs' INTO TABLE article5 FIELDS TERMINATED BY ',' ENCLOSED BY '"' ;
SELECT SUM(kingaku) FROM article6 WHERE day BETWEEN '2022-01-11' AND '2022-01-25';
SELECT noukamei, SUM(kingaku) FROM article6 WHERE day BETWEEN '2022-01-11' AND '2022-01-25' GROUP BY noukamei ORDER BY kumikan;
SELECT * FROM article6 WHERE day BETWEEN '2022-01-11' AND '2022-01-25' AND kumikan = '12';
SELECT number, byoumei, kingaku FROM article6 WHERE day BETWEEN '2022-01-11' AND '2022-01-25' AND kumikan = '12';
SELECT number, byoumei, kingaku FROM article6 WHERE day BETWEEN '2022-01-11' AND '2022-01-25' AND kumikan = '12' INTO OUTFILE 'C:/ProgramData/MySQL/MySQL Server 8.0/Uploads/sa.csv';

 SELECT number, byoumei, kingaku FROM article6 WHERE d
 SELECT number, byoumei, kingaku FROM article6 WHERE day BETWEEN '2022-01-11' AND '2022-01-25' AND kumikan = '12' INTO OUTFILE 'C:/ProgramData/MySQL/MySQL Server 8.0/Uploads/sa2.csv' CHARACTER SET 'sjis' FIELDS TERMINATED BY ',' OPTIONALLY ENCLOSED BY '"';
 SELECT 農家名, SUM(金額) FROM article6 WHERE day BETWEEN '2022-01-11' AND '2022-01-25' GROUP BY 農家名 ORDER BY 組勘番号 INTO OUTFILE 'C:/ProgramData/MySQL/MySQL Server 8.0/Uploads/goukei.csv' CHARACTER SET 'sjis' FIELDS TERMINATED BY ',' OPTIONALLY ENCLOSED BY '"';
  SELECT number, 病名, 金額 FROM article6 WHERE day  
  SELECT DABETWEEN '2022-01-11' AND '2022-01-25' AND 組勘番号 = '12';TE_FORMAT(day, '%Y%m') as YM, sum(金額) FROM article6 WHERE 農家名 = '田中和也' GROUP by DATE_FORMAT(day, '%Y%m');
  SELECT DATE_FORMAT(day, '%Y') as Y, sum(金額) FROM article6 WHERE 農家名 = '田中和也' GROUP by DATE_FORMAT(day, '%Y');
  SELECT 日時, 損坊内容, 金額 FROM article7 WHERE 日時 BETWEEN '2021-11-30' AND '2022-01-31' AND 組勘番号 = '12'; 
  SELECT 農家名,SUM(金額) FROM article7 WHERE 日時 BETWEEN '2021-11-30' AND '2022-01-31' GROUP BY 農家名 ORDER BY 組勘番号;
   SELECT * FROM table4 WHERE personalID = '1417018169'INTO OUTFILE 'C:/ProgramData/MySQL/MySQL Server 8.0/Uploads/1121.csv' CHARACTER SET 'sjis' FIELDS TERMINATED BY ',' OPTIONALLY ENCLOSED BY '"';
   SELECT DATE_FORMAT(day, '%Y') as Y, sum(金額), 農家名  FROM article6 GROUP by DATE_FORMAT(day, '%Y'), 農家名;
   SELECT DATE_FORMAT(日時, '%Y') as Y, sum(金額), 農家名  FROM article7 GROUP by DATE_FORMAT(日時, '%Y'), 農家名
   SELECT DATE_FORMAT(日時, '%Y') as Y, sum(金額) FROM article7 GROUP by DATE_FORMAT(日時, '%Y');
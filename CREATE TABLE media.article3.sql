CREATE TABLE media.article3 (
    id INT AUTO_INCREMENT NOT NULL PRIMARY KEY,
    kumikanbangou VARCHAR(50),
    noukamei  TEXT,
    number INT,
    byoumei TEXT,
    kingaku DECIMAL(10,2),
    day DATE,
    create_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP
);
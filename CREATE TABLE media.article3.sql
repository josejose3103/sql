CREATE TABLE media.article7 (
    id INT AUTO_INCREMENT NOT NULL PRIMARY KEY,
    kumikanbangou VARCHAR(50),
    noukamei  TEXT,
    byoumei TEXT,
    kingaku DECIMAL(10,2),
    day DATE,
    create_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP
);
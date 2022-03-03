CREATE TABLE media.article7 (
    id INT AUTO_INCREMENT NOT NULL PRIMARY KEY,
    day DATE,
    noukamei  TEXT,
    byoumei TEXT,
    kingaku DECIMAL(10,2),
    kumikanbangou VARCHAR(50),
    create_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP
);
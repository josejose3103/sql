CREATE TABLE media.article5 (
    id INT AUTO_INCREMENT NOT NULL PRIMARY KEY,
    kumikan int,
    noukamei VARCHAR(50),
    byoumei  VARCHAR(50),
    kingaku DECIMAL(10,2),
    day DATE,
    create_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP
);


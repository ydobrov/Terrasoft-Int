CREATE DATABASE dbname;
USE dbname;
CREATE TABLE IF NOT EXISTS LogServico (
    id INT AUTO_INCREMENT,
    Date DATETIMe NULL,
    Thread TEXT NULL,
    Level TEXT NULL,
    Logger TEXT NULL,
    Message TEXT NULL,
    Exception TEXT NULL,
    PRIMARY KEY (id)
)  ENGINE=INNODB;
CREATE DATABASE etl_db;
USE etl_db;

CREATE TABLE demographic(
id INT AUTO_INCREMENT PRIMARY KEY,
State VARCHAR(20),
TotalPop FLOAT,
Hispanic_perc FLOAT,
White_perc FLOAT,
Black_perc FLOAT,
Native_perc FLOAT,
Asian_perc FLOAT,
Pacific_perc FLOAT);

CREATE TABLE people_shot(
id INT AUTO_INCREMENT PRIMARY KEY,
State VARCHAR(20),
Number_of_People INT);




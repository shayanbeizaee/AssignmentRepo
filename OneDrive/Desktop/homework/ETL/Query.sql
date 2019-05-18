USE etl_db;

SELECT * FROM demographic;
SELECT * FROM people_shot;

CREATE TABLE analysis(
SELECT d.*, p.Number_of_People FROM 
demographic d
JOIN people_shot p 
ON (d.State = p.State)
);

SELECT * FROM analysis;

/*creating database and tables*/
CREATE DATABASE ORG;
SHOW DATABASES;
USE ORG;

CREATE TABLE Worker (
	WORKER_ID INT NOT NULL PRIMARY KEY AUTO_INCREMENT,
	FIRST_NAME CHAR(25),
	LAST_NAME CHAR(25),
	SALARY INT(15),
	JOINING_DATE DATETIME,
	DEPARTMENT CHAR(25)
);

INSERT INTO Worker 
	(WORKER_ID, FIRST_NAME, LAST_NAME, SALARY, JOINING_DATE, DEPARTMENT) VALUES
		(001, 'Monika', 'Arora', 100000, '14-02-20 09.00.00', 'HR'),
		(002, 'Niharika', 'Verma', 80000, '14-06-11 09.00.00', 'Admin'),
		(003, 'Vishal', 'Singhal', 300000, '14-02-20 09.00.00', 'HR'),
		(004, 'Amitabh', 'Singh', 500000, '14-02-20 09.00.00', 'Admin'),
		(005, 'Vivek', 'Bhati', 500000, '14-06-11 09.00.00', 'Admin'),
		(006, 'Vipul', 'Diwan', 200000, '14-06-11 09.00.00', 'Account'),
		(007, 'Satish', 'Kumar', 75000, '14-01-20 09.00.00', 'Account'),
		(008, 'Geetika', 'Chauhan', 90000, '14-04-11 09.00.00', 'Admin');

CREATE TABLE Bonus (
	WORKER_REF_ID INT,
	BONUS_AMOUNT INT(10),
	BONUS_DATE DATETIME,
	FOREIGN KEY (WORKER_REF_ID)
		REFERENCES Worker(WORKER_ID)
        ON DELETE CASCADE
);

INSERT INTO Bonus 
	(WORKER_REF_ID, BONUS_AMOUNT, BONUS_DATE) VALUES
		(001, 5000, '16-02-20'),
		(002, 3000, '16-06-11'),
		(003, 4000, '16-02-20'),
		(001, 4500, '16-02-20'),
		(002, 3500, '16-06-11');
		
CREATE TABLE Title (
	WORKER_REF_ID INT,
	WORKER_TITLE CHAR(25),
	AFFECTED_FROM DATETIME,
	FOREIGN KEY (WORKER_REF_ID)
		REFERENCES Worker(WORKER_ID)
        ON DELETE CASCADE
);

INSERT INTO Title 
	(WORKER_REF_ID, WORKER_TITLE, AFFECTED_FROM) VALUES
 (001, 'Manager', '2016-02-20 00:00:00'),
 (002, 'Executive', '2016-06-11 00:00:00'),
 (008, 'Executive', '2016-06-11 00:00:00'),
 (005, 'Manager', '2016-06-11 00:00:00'),
 (004, 'Asst. Manager', '2016-06-11 00:00:00'),
 (007, 'Executive', '2016-06-11 00:00:00'),
 (006, 'Lead', '2016-06-11 00:00:00'),
 (003, 'Lead', '2016-06-11 00:00:00');
 
 
 
/*now producing sql commands from the website for practice*/
/* website: https://www.techbeamers.com/sql-query-questions-answers-for-practice/ */


SELECT FIRST_NAME AS WORKER_NAME
FROM Worker;

SELECT UPPER(FIRST_NAME)
FROM Worker;

SELECT DISTINCT(DEPARTMENT)
FROM Worker;

SELECT SUBSTRING(FIRST_NAME,1,3)
FROM Worker;

/*finding the position of a character */
SELECT INSTR(FIRST_NAME, BINARY'a')
FROM Worker
WHERE FIRST_NAME = 'Amitabh'

SELECT RTRIM(FIRST_NAME)
FROM Worker;

SELECT LTRIM(DEPARTMENT)
FROM Worker;

SELECT DISTINCT(DEPARTMENT), COUNT(DEPARTMENT)
FROM Worker;

SELECT REPLACE(FIRST_NAME,'a','A')
FROM Worker;

SELECT CONCAT(FIRST_NAME + ' ' + LAST_NAME) AS COMPLETE_NAME
FROM Worker;

SELECT * FROM Worker
ORDER BY FIRST_NAME ASC;

SELECT * FROM Worker
ORDER BY FIRST_NAME ASC, DEPARTMENT DESC;

SELECT * FROM Worker
WHERE (FIRST_NAME = 'Vipul' AND FIRST_NAME = 'Satish');

SELECT * FROM Worker
WHERE FIRST_NAME NOT IN 'Vipul' AND 'Satish';

SELECT * FROM Worker
WHERE DEPARTMENT = 'Admin';

SELECT * FROM Worker
WHERE FIRST_NAME IS LIKE '%a%';

SELECT * FROM Worker
WHERE FIRST_NAME LIKE '_____h';

SELECT * FROM Worker
WHERE SALARY >= 100000 AND SALARY <= 500000;

/*working with date-time format*/
SELECT * FROM Worker
WHERE YEAR(JOINING_DATE) = 2014 AND MONTH(JOINING_DATE)='Februrary';

SELECT COUNT(DEPARTMENT)
FROM Worker
WHERE DEPARTMENT = 'Admin';

/*fetching data when we have to deal with index and finding the max,min or in range*/
SELECT CONCAT(FIRST_NAME + ' ' + LAST_NAME) AS WORKER_NAME, SALARY
FROM Worker
WHERE WORKER_ID IN
	(SELECT SALARY >= 100000 AND SALARY <=500000);

SELECT COUNT(WORKER_ID), DEPARTMENT
FROM Worker
GROUP BY DEPARTMENT
ORDER BY WORKER_ID DESC;

SELECT * FROM Worker
WHERE WORKER_TITLE = 'Manager';

/* fetching duplicate records*/
SELECT WORKER_TITLE, AFFECTED_FROM, COUNT(*)
FROM Title
GROUPBY WORKER_TITLE, AFFECTED_FROM
HAVING COUNT(*) >1;

/*selecting odd and even rows in sql*/
/*selecting even rows*/
SELECT * FROM Worker
WHERE MOD(id,2) = 0;

/*selecting odd rows*/
SELECT * FROM Worker
WHERE MOD(id,2) <> 0;

/*Fetching records which intersects*/
(SELECT * FROM Worker) 
INTERSECT
(SELECT * FROM WorkerClone)

/*fetching records from a table where those records are not in another table*/
(SELECT * FROM Worker)
MINUS
(SELECT * FROM Title);


	


USE GymDatabase;

CREATE TABLE Goal(
Id int primary key identity(0,1),
Grade varchar(10),
Goal int DEFAULT 0
);

--Insert boulders
INSERT INTO Goal(Grade)
VALUES('V0');
INSERT INTO Goal(Grade)
VALUES('V1');
INSERT INTO Goal(Grade)
VALUES('V2');
INSERT INTO Goal(Grade)
VALUES('V3');
INSERT INTO Goal(Grade)
VALUES('V4');
INSERT INTO Goal(Grade)
VALUES('V5');
INSERT INTO Goal(Grade)
VALUES('V6');
INSERT INTO Goal(Grade)
VALUES('V7');
INSERT INTO Goal(Grade)
VALUES('V8');
INSERT INTO Goal(Grade)
VALUES('V9');

--Insert ropes
INSERT INTO Goal(Grade)
VALUES('5.6');
INSERT INTO Goal(Grade)
VALUES('5.7');
INSERT INTO Goal(Grade)
VALUES('5.8');
INSERT INTO Goal(Grade)
VALUES('5.9');
INSERT INTO Goal(Grade)
VALUES('5.10-');
INSERT INTO Goal(Grade)
VALUES('5.10+');
INSERT INTO Goal(Grade)
VALUES('5.11-');
INSERT INTO Goal(Grade)
VALUES('5.11+');
INSERT INTO Goal(Grade)
VALUES('5.12-');
INSERT INTO Goal(Grade)
VALUES('5.12+');
INSERT INTO Goal(Grade)
VALUES('5.13-');
INSERT INTO Goal(Grade)
VALUES('5.13+');
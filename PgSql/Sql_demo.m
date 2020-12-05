
clear;
clc;
% For MySQL:
% db_name        = 'db_name';
% db_user_name   = 'db_user';
% db_user_pass   = 'db_user_pass';
% db_server_ip   = 'db_ip';
% db_server_port = '3306';
% url            = strcat('jdbc:mysql://', db_server_ip, ':', db_server_port, '/');
% dbh            = database(db_name, db_user_name, db_user_pass, 'com.mysql.jdbc.Driver', url);

% For PostgreSQL:
db_name        = 'postgres';
db_user_name   = 'postgres';
db_user_pass   = 'postgre';
db_server_ip   = 'localhost';
db_server_port = '5432';
url            = strcat('jdbc:postgresql://', db_server_ip, ':', db_server_port, '/');
connection            = database(db_name, db_user_name, db_user_pass, 'org.postgresql.Driver', url);
% %连接数据库 
% connection=database('db_name','postgres','postgres','org.postgresql.Driver','jdbc:postgresql://localhost:5432/db_name');
%% 构造SQL语句
% 创建数据库语句
create_database_command = 'CREATE DATABASE demo;';
% 删除数据库语句
delete_database_command = 'DROP DATABASE demo;';
% 创建一个表
create_table_command = 'CREATE TABLE COMPANY( ID INT PRIMARY KEY NOT NULL, NAME  TEXT NOT NULL, AGE INT NOT NULL, ADDRESS CHAR(50), SALARY REAL);';
create_table_command_0 = 'CREATE TABLE DEPARTMENT( ID INT PRIMARY KEY NOT NULL, DEPT CHAR(50) NOT NULL,  EMP_ID  INT  NOT NULL);'; 
% 删除一个表
delete_table_command = 'drop table department, company'

% postgresql模式 schema
command0= 'create schema myschema;';
command1= 'create table myschema.company( ID INT NOT NULL, NAME VARCHAR(20) NOT NULL, AGE INT NOT NULL,ADDRESS CHAR(25), SALARY DECIMAL(18,2), PRIMARY KEY(ID));';
% 查询表格是否创建
command2 = 'select * from myschema.company;'
%删除一个为空的模式
command3 = 'DROP SCHEMA myschema;'

%% PostgreSQL INSERT INTO 语句用于向表中插入新记录
command4 ='CREATE TABLE COMPANY( ID INT PRIMARY KEY NOT NULL, NAME  TEXT NOT NULL,AGE INT NOT NULL,ADDRESS CHAR(50), SALARY  REAL,JOIN_DATE DATE);'
command5 = "INSERT INTO COMPANY (ID,NAME,AGE,ADDRESS,SALARY,JOIN_DATE) VALUES (1, 'Paul', 32, 'California', 20000.00,'2001-07-13');"
command6 = "INSERT INTO COMPANY (ID,NAME,AGE,ADDRESS,JOIN_DATE) VALUES (2, 'Allen', 25, 'Texas', '2007-12-13');"
command7 = "INSERT INTO COMPANY (ID,NAME,AGE,ADDRESS,SALARY,JOIN_DATE) VALUES (3, 'Teddy', 23, 'Norway', 20000.00, DEFAULT );"
command8 = "INSERT INTO COMPANY (ID,NAME,AGE,ADDRESS,SALARY,JOIN_DATE) VALUES (4, 'Mark', 25, 'Rich-Mond ', 65000.00, '2007-12-13' ), (5, 'David', 27, 'Texas', 85000.00, '2007-12-13');"
%  SELECT 语句查询表格数据
command9 = "SELECT * FROM company;"
%% PostgreSQL SELECT 语句
command10 = "SELECT * FROM company;"
% 以下 SQL 语句将删除 ID 为 2 的数据：
command11 = " DELETE FROM COMPANY WHERE ID = 2;"
% 以下语句将删除整张 COMPANY 表：
command12 = "DELETE FROM COMPANY;"
% SELECT ID,NAME FROM company;
% 使用了布尔表达式（SALARY=10000）来查询数据：
% SELECT * FROM COMPANY WHERE SALARY = 10000
% sqlcommand='select * from postgres;'
% %执行SQL语句
cursor=exec(connection,command9);
% %获取指定数量的数据
row=fetch(cursor);
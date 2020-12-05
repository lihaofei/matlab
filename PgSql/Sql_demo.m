
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
% %�������ݿ� 
% connection=database('db_name','postgres','postgres','org.postgresql.Driver','jdbc:postgresql://localhost:5432/db_name');
%% ����SQL���
% �������ݿ����
create_database_command = 'CREATE DATABASE demo;';
% ɾ�����ݿ����
delete_database_command = 'DROP DATABASE demo;';
% ����һ����
create_table_command = 'CREATE TABLE COMPANY( ID INT PRIMARY KEY NOT NULL, NAME  TEXT NOT NULL, AGE INT NOT NULL, ADDRESS CHAR(50), SALARY REAL);';
create_table_command_0 = 'CREATE TABLE DEPARTMENT( ID INT PRIMARY KEY NOT NULL, DEPT CHAR(50) NOT NULL,  EMP_ID  INT  NOT NULL);'; 
% ɾ��һ����
delete_table_command = 'drop table department, company'

% postgresqlģʽ schema
command0= 'create schema myschema;';
command1= 'create table myschema.company( ID INT NOT NULL, NAME VARCHAR(20) NOT NULL, AGE INT NOT NULL,ADDRESS CHAR(25), SALARY DECIMAL(18,2), PRIMARY KEY(ID));';
% ��ѯ����Ƿ񴴽�
command2 = 'select * from myschema.company;'
%ɾ��һ��Ϊ�յ�ģʽ
command3 = 'DROP SCHEMA myschema;'

%% PostgreSQL INSERT INTO �����������в����¼�¼
command4 ='CREATE TABLE COMPANY( ID INT PRIMARY KEY NOT NULL, NAME  TEXT NOT NULL,AGE INT NOT NULL,ADDRESS CHAR(50), SALARY  REAL,JOIN_DATE DATE);'
command5 = "INSERT INTO COMPANY (ID,NAME,AGE,ADDRESS,SALARY,JOIN_DATE) VALUES (1, 'Paul', 32, 'California', 20000.00,'2001-07-13');"
command6 = "INSERT INTO COMPANY (ID,NAME,AGE,ADDRESS,JOIN_DATE) VALUES (2, 'Allen', 25, 'Texas', '2007-12-13');"
command7 = "INSERT INTO COMPANY (ID,NAME,AGE,ADDRESS,SALARY,JOIN_DATE) VALUES (3, 'Teddy', 23, 'Norway', 20000.00, DEFAULT );"
command8 = "INSERT INTO COMPANY (ID,NAME,AGE,ADDRESS,SALARY,JOIN_DATE) VALUES (4, 'Mark', 25, 'Rich-Mond ', 65000.00, '2007-12-13' ), (5, 'David', 27, 'Texas', 85000.00, '2007-12-13');"
%  SELECT ����ѯ�������
command9 = "SELECT * FROM company;"
%% PostgreSQL SELECT ���
command10 = "SELECT * FROM company;"
% ���� SQL ��佫ɾ�� ID Ϊ 2 �����ݣ�
command11 = " DELETE FROM COMPANY WHERE ID = 2;"
% ������佫ɾ������ COMPANY ��
command12 = "DELETE FROM COMPANY;"
% SELECT ID,NAME FROM company;
% ʹ���˲������ʽ��SALARY=10000������ѯ���ݣ�
% SELECT * FROM COMPANY WHERE SALARY = 10000
% sqlcommand='select * from postgres;'
% %ִ��SQL���
cursor=exec(connection,command9);
% %��ȡָ������������
row=fetch(cursor);
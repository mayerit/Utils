SET DATEFORMAT DMY

insert into teste (strvalue) values ('...teste')
select * from teste

select * from func
insert into func (teste, idemployee) values ('Luiz Fernando', 1)

select * from depend
insert into depend()


insert into employee (firstname, lastname, designation, intvalue, decvalue, datevalue)
values ('luiz fernando', 'mayer', 'programador', 123, 123.59, '09/10/2015 15:15:15.555')
insert into employee (firstname, lastname, designation, intvalue, decvalue, datevalue)
values ('luiz fernando', 'mayer', 'programador', 123, 123456789.99, '09/10/2015 15:15:15.555')
insert into employee (firstname, lastname, designation, intvalue, decvalue, datevalue)
values ('luiz fernando', 'mayer', 'programador', 123, 123456789.99, GETDATE())

update employee set firstname = 'alterado', datevalue = GETDATE() where id = 1 

select count(*) from employee
select * from employee order by id DESC
delete from employee
select SCOPE_IDENTITY()
select * from employee where id = 4037



ALTER TABLE Employee ADD UNIQUE (guid)



declare 
	@teste varchar(50),
	@dtteste date,
	@mn money,
	@intValue int 
SET @teste = 'xinxa'
SET @dtteste = getdate()
set @mn = 12345.67
print @dtteste
select @teste, @dtteste, 'R$' + cast(@mn as varchar(50)), @mn + 1000
SET @intvalue = 1
if @intvalue = 1 
begin
	Set @intvalue = 2 
end;
else
begin 
	Set @intvalue = 2 
end;
select @teste, @dtteste, 'R$' + cast(@mn as varchar(50)), @mn + 1000, @intvalue
--**************************************************
while (@intvalue <=10)
begin
	print('**** valor ****' + cast(@intvalue as varchar(50)))
	SET @intvalue = @intvalue + 1
end;




-- Transaction
http://stackoverflow.com/questions/506602/best-way-to-work-with-transactions-in-ms-sql-server-management-studio
--
http://sqlmag.com/t-sql/t-sql-101-lesson-1


http://www.sqlserverdicas.com/2010/12/cursores-exemplo-basico-de-utilizacao.html
http://www.codeproject.com/Articles/15222/How-to-Use-Stored-Procedures-in-VB
http://www.w3schools.com/asp/met_comm_createparameter.asp

--
http://www.databasejournal.com/features/mssql/article.php/3087431/T-SQL-Programming-Part-1---Defining-Variables-and-IFELSE-logic.htm
http://www.databasejournal.com/features/mssql/t-sql-programming-part-15-understanding-how-to-write-a-correlated-subquery

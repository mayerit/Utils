USE [DBVICERI]
GO
select * from [dbo].[Corretor]
select * from [dbo].[Cliente] order by IdCliente DESC
select * from [dbo].[ClienteCorretor]
select * from [dbo].[UF]
select * from [dbo].[Cidade] WHERE iD = 3284
update [dbo].[Cliente] set ativo = 0 where IdCliente = 30 
insert into ClienteCorretor values (2, 30)
--
delete from ClienteCorretor where IdCliente in (select IdCliente from Cliente Where cpf = '12349227898')

delete from [dbo].[ClienteCorretor]
delete from [dbo].[Corretor]
delete from [dbo].[Cliente]

select idcorretor, codigo, nome, cpf from corretor order by nome  


select cliente.IdCliente, cliente.Nome, cliente.CPF, 
Ativo = CASE cliente.Ativo
      WHEN 1 THEN 'Sim'
	  WHEN 0 THEN 'Não'   
   END ,
   cliente.Ativo, Corretor.IdCorretor, Corretor.Nome, Corretor.Codigo, cliente.CidadeID, UF.Nome as UF, d.Nome as Cidade
  from cliente 
inner join ClienteCorretor b on cliente.IdCliente = b.IdCliente
inner join Corretor  on b.IdCorretor = Corretor.IdCorretor
inner join Cidade d on cliente.CidadeID = d.ID
inner join UF on d.IDUF = UF.ID
where uf.id = 13
 
where cliente.ativo = 0 and 
	cliente.nome like 'Nome do Cliente XXXXXX%' and 
	corretor.nome like 'Nome do Corretor XXXXXX%' and 
	corretor.codigo like '1212%' and cliente.cpf = '12349227898'




delete from [dbo].[Corretor]

INSERT INTO [dbo].[Corretor] ([Codigo],[Nome],[CPF])
VALUES ('1212','Corretor 1','12349227898')
GO

INSERT INTO [dbo].[Corretor] ([Codigo],[Nome],[CPF])
VALUES ('1213','Corretor 2','13449227898')
GO

INSERT INTO [dbo].[Corretor] ([Codigo],[Nome],[CPF])
VALUES ('1214','Corretor 3','14449227898')
GO

--CLiente 
select * from [dbo].[Cliente]
INSERT INTO [dbo].[Cliente]([Nome],[CPF],[Endereco],[Ativo],[CidadeID])
     VALUES ('Luiz Fernando Martins Mayer', '12349227898', 'Rua das Campânulas, 27', 1, 3284)
GO
INSERT INTO [dbo].[Cliente]([Nome],[CPF],[Endereco],[Ativo],[CidadeID])
     VALUES ('Cliente 2****', '22349227898', 'Rua das Campânulas, 27', 1, 3284)
GO
INSERT INTO [dbo].[Cliente]([Nome],[CPF],[Endereco],[Ativo],[CidadeID])
     VALUES ('Cliente 3****', '32349227898', 'Rua das Campânulas, 27', 1, 3284)
GO

--ClienteCorretor
select * from [dbo].[ClienteCorretor]
INSERT INTO [dbo].[ClienteCorretor] values (1, 1)
GO
INSERT INTO [dbo].[ClienteCorretor] values (2, 2)
GO
INSERT INTO [dbo].[ClienteCorretor] values (2, 3)
GO



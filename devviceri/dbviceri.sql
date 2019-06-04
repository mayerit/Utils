-- ****************** SqlDBM: Microsoft SQL Server ******************
-- ******************************************************************
DROP TABLE [dbo].[ClienteCorretor]
GO
DROP TABLE [dbo].[Cliente]
GO
DROP TABLE [dbo].[Corretor]
GO
DROP TABLE [dbo].[Cidade]
GO
DROP TABLE [dbo].[UF]
GO
-- ************************************** [dbo].[UF]
CREATE TABLE [dbo].[UF]
(
 [ID]   int NOT NULL ,
 [Nome] varchar(50) NOT NULL ,
 CONSTRAINT [PK_UF] PRIMARY KEY CLUSTERED ([ID] ASC)
);
GO
-- ************************************** [dbo].[Cidade]

CREATE TABLE [dbo].[Cidade]
(
 [ID]   int NOT NULL ,
 [Nome] varchar(50) NOT NULL ,
 [IDUF] int NOT NULL ,
 CONSTRAINT [PK_Cidade] PRIMARY KEY CLUSTERED ([ID] ASC),
 CONSTRAINT [FK_24] FOREIGN KEY ([IDUF])  REFERENCES [dbo].[UF]([ID])
);
GO

CREATE NONCLUSTERED INDEX [fkIdx_24] ON [dbo].[Cidade] 
 (
  [IDUF] ASC
 )
GO
-- ************************************** [dbo].[Corretor]
 
CREATE TABLE [dbo].[Corretor]
(
 [IdCorretor] int IDENTITY NOT NULL ,
 [Codigo]     varchar(50) NOT NULL ,
 [Nome]       varchar(50) NOT NULL ,
 [CPF]        varchar(50) NOT NULL ,
 CONSTRAINT [PK_Corretor] PRIMARY KEY CLUSTERED ([IdCorretor] ASC)
);
GO
-- ************************************** [dbo].[Cliente]

CREATE TABLE [dbo].[Cliente]
(
 [IdCliente] int IDENTITY NOT NULL ,
 [Nome]      varchar(50) NOT NULL ,
 [CPF]       varchar(50) NOT NULL ,
 [Endereco]  varchar(50) NOT NULL ,
 [Ativo]     bit NOT NULL ,
 [CidadeID]        int NOT NULL ,
 CONSTRAINT [PK_Cliente] PRIMARY KEY CLUSTERED ([IdCliente] ASC),
 CONSTRAINT [FK_33] FOREIGN KEY ([CidadeID])  REFERENCES [dbo].[Cidade]([ID])
);
GO

CREATE NONCLUSTERED INDEX [fkIdx_33] ON [dbo].[Cliente] 
 (
  [CidadeID] ASC
 )
GO

-- ************************************** [dbo].[ClienteCorretor]
CREATE TABLE [dbo].[ClienteCorretor]
(
 [IdCorretor] int NOT NULL ,
 [IdCliente]  int NOT NULL ,
 CONSTRAINT [PK_ClienteCorretor] PRIMARY KEY CLUSTERED ([IdCorretor] ASC, [IdCliente] ASC),
 CONSTRAINT [FK_44] FOREIGN KEY ([IdCorretor])  REFERENCES [dbo].[Corretor]([IdCorretor]),
 CONSTRAINT [FK_48] FOREIGN KEY ([IdCliente])  REFERENCES [dbo].[Cliente]([IdCliente])
);
GO

CREATE NONCLUSTERED INDEX [fkIdx_44] ON [dbo].[ClienteCorretor] 
 (
  [IdCorretor] ASC
 )
GO

CREATE NONCLUSTERED INDEX [fkIdx_48] ON [dbo].[ClienteCorretor] 
 (
  [IdCliente] ASC
 )
GO

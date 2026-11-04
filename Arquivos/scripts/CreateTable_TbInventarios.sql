-- TbInventarios: cabecalho do inventario (SPED Bloco H).
-- Idempotente: cria a tabela se nao existir e acrescenta so as colunas que faltarem
-- (cliente sem nada, com a tabela completa ou com uma versao antiga/parcial).
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO

IF OBJECT_ID('dbo.TbInventarios', 'U') IS NULL
BEGIN
CREATE TABLE [dbo].[TbInventarios](
	[IdInventario] [int] NOT NULL,
	[IdFilial] [smallint] NOT NULL,
	[Data] [datetime] NULL,
	[IDFuncionario] [smallint] NOT NULL,
	[AnoInventario] [varchar](4) NOT NULL,
	[PeriodoInventarioTipo] [varchar](10) NOT NULL,
	[PeriodoInventario] [varchar](2) NOT NULL,
	[DataInventario] [datetime] NULL,
	[Fechado] [bit] NOT NULL,
	[Nada] [varchar](1) NOT NULL,
	[quick~rs] [timestamp] NULL,
 CONSTRAINT [IdInventario] PRIMARY KEY CLUSTERED 
(
	[IdInventario] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, FILLFACTOR = 90) ON [PRIMARY]
) ON [PRIMARY]
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.TbInventarios') AND name = 'IdInventario')
    ALTER TABLE dbo.TbInventarios ADD [IdInventario] [int] NOT NULL DEFAULT (0)
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.TbInventarios') AND name = 'IdFilial')
    ALTER TABLE dbo.TbInventarios ADD [IdFilial] [smallint] NOT NULL DEFAULT (0)
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.TbInventarios') AND name = 'IDFuncionario')
    ALTER TABLE dbo.TbInventarios ADD [IDFuncionario] [smallint] NOT NULL DEFAULT (0)
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.TbInventarios') AND name = 'AnoInventario')
    ALTER TABLE dbo.TbInventarios ADD [AnoInventario] [varchar](4) NOT NULL DEFAULT ('')
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.TbInventarios') AND name = 'PeriodoInventarioTipo')
    ALTER TABLE dbo.TbInventarios ADD [PeriodoInventarioTipo] [varchar](10) NOT NULL DEFAULT ('')
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.TbInventarios') AND name = 'PeriodoInventario')
    ALTER TABLE dbo.TbInventarios ADD [PeriodoInventario] [varchar](2) NOT NULL DEFAULT ('')
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.TbInventarios') AND name = 'Fechado')
    ALTER TABLE dbo.TbInventarios ADD [Fechado] [bit] NOT NULL DEFAULT (0)
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.TbInventarios') AND name = 'Nada')
    ALTER TABLE dbo.TbInventarios ADD [Nada] [varchar](1) NOT NULL DEFAULT ('')
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.TbInventarios') AND name = 'Data')
    ALTER TABLE dbo.TbInventarios ADD [Data] [datetime] NULL
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.TbInventarios') AND name = 'DataInventario')
    ALTER TABLE dbo.TbInventarios ADD [DataInventario] [datetime] NULL
GO

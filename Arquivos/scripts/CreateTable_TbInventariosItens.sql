-- TbInventariosItens: itens do inventario (SPED Bloco H). Depende da TbInventarios.
-- Idempotente: cria a tabela se nao existir e acrescenta so as colunas que faltarem
-- (cliente sem nada, com a tabela completa ou com uma versao antiga/parcial).
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO

IF OBJECT_ID('dbo.TbInventariosItens', 'U') IS NULL
BEGIN
CREATE TABLE [dbo].[TbInventariosItens](
	[IdInventario] [int] NOT NULL,
	[Seq] [int] NOT NULL,
	[IDProduto] [int] NOT NULL,
	[NomeProduto] [varchar](220) NOT NULL,
	[SaldoCalculado] [decimal](12, 5) NOT NULL,
	[VlrUnitInvent] [decimal](11, 3) NOT NULL,
	[TotalInvent] [decimal](12, 2) NOT NULL,
	[MetaCalculado] [decimal](12, 5) NOT NULL,
	[SaldoLancado] [decimal](12, 5) NOT NULL,
	[quick~rs] [timestamp] NULL,
 CONSTRAINT [IdInventarioItens] PRIMARY KEY CLUSTERED 
(
	[IdInventario] ASC,
	[Seq] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, FILLFACTOR = 90) ON [PRIMARY]
) ON [PRIMARY]
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.TbInventariosItens') AND name = 'IdInventario')
    ALTER TABLE dbo.TbInventariosItens ADD [IdInventario] [int] NOT NULL DEFAULT (0)
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.TbInventariosItens') AND name = 'Seq')
    ALTER TABLE dbo.TbInventariosItens ADD [Seq] [int] NOT NULL DEFAULT (0)
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.TbInventariosItens') AND name = 'IDProduto')
    ALTER TABLE dbo.TbInventariosItens ADD [IDProduto] [int] NOT NULL DEFAULT (0)
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.TbInventariosItens') AND name = 'NomeProduto')
    ALTER TABLE dbo.TbInventariosItens ADD [NomeProduto] [varchar](220) NOT NULL DEFAULT ('')
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.TbInventariosItens') AND name = 'SaldoCalculado')
    ALTER TABLE dbo.TbInventariosItens ADD [SaldoCalculado] [decimal](12, 5) NOT NULL DEFAULT (0)
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.TbInventariosItens') AND name = 'VlrUnitInvent')
    ALTER TABLE dbo.TbInventariosItens ADD [VlrUnitInvent] [decimal](11, 3) NOT NULL DEFAULT (0)
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.TbInventariosItens') AND name = 'TotalInvent')
    ALTER TABLE dbo.TbInventariosItens ADD [TotalInvent] [decimal](12, 2) NOT NULL DEFAULT (0)
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.TbInventariosItens') AND name = 'MetaCalculado')
    ALTER TABLE dbo.TbInventariosItens ADD [MetaCalculado] [decimal](12, 5) NOT NULL DEFAULT (0)
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.TbInventariosItens') AND name = 'SaldoLancado')
    ALTER TABLE dbo.TbInventariosItens ADD [SaldoLancado] [decimal](12, 5) NOT NULL DEFAULT (0)
GO

IF NOT EXISTS (SELECT 1 FROM sys.foreign_keys WHERE name = 'FK_TbInventariosItens_TbInventarios')
    ALTER TABLE dbo.TbInventariosItens WITH NOCHECK ADD CONSTRAINT FK_TbInventariosItens_TbInventarios
        FOREIGN KEY (IdInventario) REFERENCES dbo.TbInventarios (IdInventario) ON UPDATE CASCADE ON DELETE CASCADE
GO

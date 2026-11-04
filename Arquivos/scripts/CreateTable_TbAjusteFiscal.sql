-- TbAjusteFiscal: ajustes manuais do ESTOQUE FISCAL (movimento que entra em vMovimentoFiscal como 'AJUSTE').
-- Gravada por EstoqueFiscalAjustar quando alguem altera o estoque fiscal de um produto na mao.
-- Idempotente: cria a tabela se nao existir.
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF OBJECT_ID('dbo.TbAjusteFiscal', 'U') IS NULL
BEGIN
    CREATE TABLE dbo.TbAjusteFiscal(
        IdAjuste   INT IDENTITY(1,1) NOT NULL CONSTRAINT PK_TbAjusteFiscal PRIMARY KEY,
        Data       DATETIME NOT NULL CONSTRAINT DF_TbAjusteFiscal_Data DEFAULT (GETDATE()),
        IdProduto  INT NOT NULL,
        Quantidade DECIMAL(18,3) NOT NULL,
        IdUsuario  INT NULL,
        Obs        VARCHAR(100) NULL
    )
END
GO
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_TbAjusteFiscal_Produto' AND object_id = OBJECT_ID('dbo.TbAjusteFiscal'))
    CREATE INDEX IX_TbAjusteFiscal_Produto ON dbo.TbAjusteFiscal (IdProduto, Data)
GO

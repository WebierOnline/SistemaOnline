/***************************
DESCRICAO.: Historico das correcoes de cadastro feitas pela conferencia do SPED Fiscal (tela Fiscal > SPED Fiscal,
            SpedGerador.exe): codigo do municipio e IE do fornecedor pelo XML da nota de entrada, NCM do produto pelo
            ultimo XML. Guarda o valor anterior para poder desfazer.
DATA......: 2026-11
***************************/
IF OBJECT_ID('dbo.SPED_CorrecaoCadastro') IS NULL
BEGIN
    CREATE TABLE dbo.SPED_CorrecaoCadastro (
        Id            INT IDENTITY(1,1) NOT NULL CONSTRAINT PK_SPED_CorrecaoCadastro PRIMARY KEY,
        Data          DATETIME      NOT NULL CONSTRAINT DF_SPED_CorrecaoCadastro_Data DEFAULT (GETDATE()),
        Tabela        VARCHAR(30)   NOT NULL,
        Codigo        VARCHAR(30)   NOT NULL,
        Campo         VARCHAR(30)   NOT NULL,
        ValorAnterior VARCHAR(100)  NULL,
        ValorNovo     VARCHAR(100)  NULL,
        Fonte         VARCHAR(200)  NULL,
        IdUsuario     INT           NULL
    )
    CREATE INDEX IX_SPED_CorrecaoCadastro_Tabela ON dbo.SPED_CorrecaoCadastro (Tabela, Codigo)
END
GO

-- Categorias.Padrao: marca a categoria padrao daquele Tipo_Empresa.
-- Usado por Produtos_Cadastro (produto novo ja vem com a categoria padrao selecionada)
-- e definido em Categorias_Cadastro pelo botao "Padrao".
-- A regra "so uma padrao por Tipo_Empresa" e garantida pela aplicacao (o UPDATE zera as
-- demais antes de marcar a nova), NAO por indice filtrado - indice filtrado quebra
-- INSERT/UPDATE via Provider=SQLOLEDB no VB6 (ARITHABORT OFF).
-- Idempotente.
SET NOCOUNT ON;
GO

IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('Categorias') AND name = 'Padrao'
)
    ALTER TABLE Categorias
        ADD Padrao BIT NOT NULL CONSTRAINT DF_Categorias_Padrao DEFAULT (0);
GO

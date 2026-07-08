-- ============================================================
-- Script: AlterTable_produtos_entrada_rename_addfields.sql
-- Banco  : SQL Server 2014
-- Tabela : produtos_entrada
-- Objetivo:
--   1) Renomear campos que possuem o mesmo nome que em EntradaEstoque
--   2) Adicionar campos presentes em EntradaEstoque que estavam faltando
-- ============================================================

-- ============================================================
-- PARTE 1: Renomear campos com mesmo nome que EntradaEstoque
-- ============================================================

-- DATA_EMISSAO -> DataEmissao
IF EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('produtos_entrada') AND name = 'DATA_EMISSAO'
)
    EXEC sp_rename 'produtos_entrada.DATA_EMISSAO', 'DataEmissao', 'COLUMN';
GO

-- DATA_SAIDA -> DataSaida
IF EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('produtos_entrada') AND name = 'DATA_SAIDA'
)
    EXEC sp_rename 'produtos_entrada.DATA_SAIDA', 'DataSaida', 'COLUMN';
GO

-- HORA_SAIDA -> HoraSaida
IF EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('produtos_entrada') AND name = 'HORA_SAIDA'
)
    EXEC sp_rename 'produtos_entrada.HORA_SAIDA', 'HoraSaida', 'COLUMN';
GO

-- ============================================================
-- PARTE 2: Adicionar campos faltantes (baseados em EntradaEstoque)
-- ============================================================

IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('produtos_entrada') AND name = 'CodigoNota'
)
    ALTER TABLE produtos_entrada ADD CodigoNota INT NOT NULL DEFAULT 0;
GO

IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('produtos_entrada') AND name = 'SerieNF'
)
    ALTER TABLE produtos_entrada ADD SerieNF SMALLINT NOT NULL DEFAULT 0;
GO

IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('produtos_entrada') AND name = 'ModeloNF'
)
    ALTER TABLE produtos_entrada ADD ModeloNF VARCHAR(2) NOT NULL DEFAULT '';
GO

IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('produtos_entrada') AND name = 'NaturezaOperacao'
)
    ALTER TABLE produtos_entrada ADD NaturezaOperacao VARCHAR(60) NOT NULL DEFAULT '';
GO

IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('produtos_entrada') AND name = 'RazaoSocial'
)
    ALTER TABLE produtos_entrada ADD RazaoSocial VARCHAR(60) NOT NULL DEFAULT '';
GO

IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('produtos_entrada') AND name = 'CNPJ_CPF'
)
    ALTER TABLE produtos_entrada ADD CNPJ_CPF VARCHAR(18) NOT NULL DEFAULT '';
GO

IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('produtos_entrada') AND name = 'ValorProdutos'
)
    ALTER TABLE produtos_entrada ADD ValorProdutos DECIMAL(15,2) NOT NULL DEFAULT 0;
GO

IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('produtos_entrada') AND name = 'ValorDesconto'
)
    ALTER TABLE produtos_entrada ADD ValorDesconto DECIMAL(15,2) NOT NULL DEFAULT 0;
GO

IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('produtos_entrada') AND name = 'ValorIPI'
)
    ALTER TABLE produtos_entrada ADD ValorIPI DECIMAL(15,2) NOT NULL DEFAULT 0;
GO

IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('produtos_entrada') AND name = 'ValorICMS'
)
    ALTER TABLE produtos_entrada ADD ValorICMS DECIMAL(15,2) NOT NULL DEFAULT 0;
GO

IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('produtos_entrada') AND name = 'ValorICMSST'
)
    ALTER TABLE produtos_entrada ADD ValorICMSST DECIMAL(15,2) NOT NULL DEFAULT 0;
GO

-- ============================================================
-- Verificacao final
-- ============================================================
SELECT
    name        AS Campo,
    TYPE_NAME(user_type_id) AS Tipo,
    max_length  AS Tamanho,
    is_nullable AS Nulo
FROM sys.columns
WHERE object_id = OBJECT_ID('produtos_entrada')
ORDER BY column_id;
GO

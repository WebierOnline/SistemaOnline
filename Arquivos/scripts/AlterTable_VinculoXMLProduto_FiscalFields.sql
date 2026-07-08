-- ============================================================
-- Script: AlterTable_VinculoXMLProduto_FiscalFields.sql
-- Banco  : SQL Server 2014
-- Tabela : VinculoXMLProduto
-- Objetivo: Adicionar campos fiscais da NF-e entre uCom e IDProduto
--
-- Ordem logica desejada:
--   ID, IDFornecedor, cProd, EANEmbalagem, xProd, uCom,
--   NCM, CFOP, CST, pICMS,
--   IPICST, IPIpIPI,
--   PISCST, PISpPIS,
--   COFINSCST, COFINSpCOFINS,
--   CEST,
--   IDProduto, EANProduto, Fracionamento, CustoUnitario, DataAtualizacao
--
-- Obs: ALTER TABLE ADD inclui ao final fisicamente; a ordem logica
--      e garantida pelas queries que nomeiam colunas explicitamente.
-- Todos os tipos espelham os campos correspondentes em EntradaEstoqueItens.
-- ============================================================

IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('VinculoXMLProduto') AND name = 'NCM'
)
    ALTER TABLE VinculoXMLProduto ADD NCM VARCHAR(8) NOT NULL DEFAULT '';
GO

IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('VinculoXMLProduto') AND name = 'CFOP'
)
    ALTER TABLE VinculoXMLProduto ADD CFOP INT NOT NULL DEFAULT 0;
GO

IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('VinculoXMLProduto') AND name = 'CST'
)
    ALTER TABLE VinculoXMLProduto ADD CST VARCHAR(3) NOT NULL DEFAULT '';
GO

IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('VinculoXMLProduto') AND name = 'pICMS'
)
    ALTER TABLE VinculoXMLProduto ADD pICMS DECIMAL(5,2) NOT NULL DEFAULT 0;
GO

IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('VinculoXMLProduto') AND name = 'IPICST'
)
    ALTER TABLE VinculoXMLProduto ADD IPICST VARCHAR(3) NOT NULL DEFAULT '';
GO

IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('VinculoXMLProduto') AND name = 'IPIpIPI'
)
    ALTER TABLE VinculoXMLProduto ADD IPIpIPI DECIMAL(5,2) NOT NULL DEFAULT 0;
GO

IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('VinculoXMLProduto') AND name = 'PISCST'
)
    ALTER TABLE VinculoXMLProduto ADD PISCST VARCHAR(2) NOT NULL DEFAULT '';
GO

IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('VinculoXMLProduto') AND name = 'PISpPIS'
)
    ALTER TABLE VinculoXMLProduto ADD PISpPIS DECIMAL(7,4) NOT NULL DEFAULT 0;
GO

IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('VinculoXMLProduto') AND name = 'COFINSCST'
)
    ALTER TABLE VinculoXMLProduto ADD COFINSCST VARCHAR(2) NOT NULL DEFAULT '';
GO

IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('VinculoXMLProduto') AND name = 'COFINSpCOFINS'
)
    ALTER TABLE VinculoXMLProduto ADD COFINSpCOFINS DECIMAL(7,4) NOT NULL DEFAULT 0;
GO

IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('VinculoXMLProduto') AND name = 'CEST'
)
    ALTER TABLE VinculoXMLProduto ADD CEST VARCHAR(7) NOT NULL DEFAULT '';
GO

-- ============================================================
-- Verificacao final
-- ============================================================
SELECT
    name        AS Campo,
    TYPE_NAME(user_type_id) AS Tipo,
    max_length  AS Tamanho,
    is_nullable AS Nulo,
    object_definition(default_object_id) AS ValorDefault
FROM sys.columns
WHERE object_id = OBJECT_ID('VinculoXMLProduto')
ORDER BY column_id;
GO

-- ============================================================
-- Script: AlterTable_VinculoXMLProduto_3campos.sql
-- Banco  : SQL Server 2014
-- Tabela : VinculoXMLProduto
-- Objetivo: Adicionar QuantidadeComercial, ValorUnitarioComercializacao e UNID_MEDIDA
-- Ordem logica desejada apos adicao:
--   ID, IDFornecedor, cProd, EANEmbalagem, xProd, uCom,
--   QuantidadeComercial, ValorUnitarioComercializacao,
--   IDProduto, EANProduto, UNID_MEDIDA, Fracionamento, CustoUnitario,
--   DataAtualizacao, NCM, CFOP, CST, pICMS, IPICST, IPIpIPI,
--   PISCST, PISpPIS, COFINSCST, COFINSpCOFINS, CEST
-- ============================================================

IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('VinculoXMLProduto') AND name = 'QuantidadeComercial'
)
    ALTER TABLE VinculoXMLProduto ADD QuantidadeComercial DECIMAL(9,3) NOT NULL DEFAULT 0;
GO

IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('VinculoXMLProduto') AND name = 'ValorUnitarioComercializacao'
)
    ALTER TABLE VinculoXMLProduto ADD ValorUnitarioComercializacao DECIMAL(15,4) NOT NULL DEFAULT 0;
GO

IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('VinculoXMLProduto') AND name = 'UNID_MEDIDA'
)
    ALTER TABLE VinculoXMLProduto ADD UNID_MEDIDA VARCHAR(6) NOT NULL DEFAULT '';
GO

-- ============================================================
-- Verificacao
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

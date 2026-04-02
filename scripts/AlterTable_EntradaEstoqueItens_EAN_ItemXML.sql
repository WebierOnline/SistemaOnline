-- ============================================================
-- Script: AlterTable_EntradaEstoqueItens_EAN_ItemXML.sql
-- Banco  : SQL Server 2014
-- Tabela : EntradaEstoqueItens
-- Objetivo: Adicionar os campos EAN e ItemXML utilizados na
--           importacao de XML e no formulario de vinculacao.
--
-- EAN     : GTIN/EAN da embalagem do fornecedor (<cEAN> da NF-e)
-- ItemXML : Numero do item no XML da NF-e (atributo nItem de <det>)
-- ============================================================

-- EAN: Codigo GTIN/EAN da embalagem conforme informado na NF-e.
-- Tag XML: <cEAN> dentro de <prod>
IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('EntradaEstoqueItens') AND name = 'EAN'
)
    ALTER TABLE EntradaEstoqueItens ADD EAN VARCHAR(14) NOT NULL DEFAULT '';
GO

-- ItemXML: Numero sequencial do item na NF-e (atributo nItem do elemento <det>).
-- Ex: <det nItem="1">, <det nItem="2">, etc.
IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('EntradaEstoqueItens') AND name = 'ItemXML'
)
    ALTER TABLE EntradaEstoqueItens ADD ItemXML INT NOT NULL DEFAULT 0;
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
WHERE object_id = OBJECT_ID('EntradaEstoqueItens')
  AND name IN ('EAN', 'ItemXML')
ORDER BY name;
GO

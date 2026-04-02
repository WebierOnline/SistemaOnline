-- ============================================================
-- Script: AlterTable_EntradaEstoqueItens_vItem.sql
-- Banco  : SQL Server 2014
-- Tabela : EntradaEstoqueItens
-- Objetivo: Adicionar o campo vItem para armazenar o valor
--           total liquido do item da NF-e.
--
-- vItem : Valor total do item (vProd - descontos + acrescimos
--         + tributos inclusos no preco).
--         Tag XML: <vItem> diretamente dentro de <det>
--         (elemento irmao de <prod> e <imposto>).
--         Presente apenas em NF-e com Reforma Tributaria.
-- ============================================================

IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('EntradaEstoqueItens') AND name = 'vItem'
)
    ALTER TABLE EntradaEstoqueItens ADD vItem DECIMAL(15,4) NOT NULL DEFAULT 0;
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
  AND name = 'vItem'
ORDER BY name;
GO

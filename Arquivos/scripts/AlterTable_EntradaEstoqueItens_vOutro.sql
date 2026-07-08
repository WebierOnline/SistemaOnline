-- ============================================================
-- Script: AlterTable_EntradaEstoqueItens_vOutro.sql
-- Banco  : SQL Server 2014
-- Tabela : EntradaEstoqueItens
-- Objetivo: Adicionar o campo vOutro para armazenar o valor de
--           outras despesas acessorias por item da NF-e.
--
-- vOutro : Valor de outras despesas acessorias do item.
--          Tag XML: <vOutro> dentro de <prod> (tag opcional).
--          Nao confundir com <vOutro> em <ICMSTot> (total da nota),
--          que ja e salvo em EntradaEstoque.ValorOutrasDespesas.
-- ============================================================

IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('EntradaEstoqueItens') AND name = 'vOutro'
)
    ALTER TABLE EntradaEstoqueItens ADD vOutro DECIMAL(15,4) NOT NULL DEFAULT 0;
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
  AND name = 'vOutro'
ORDER BY name;
GO

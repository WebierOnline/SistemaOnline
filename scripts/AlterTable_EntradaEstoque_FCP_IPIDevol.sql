-- ============================================================
-- Script: AlterTable_EntradaEstoque_FCP_IPIDevol.sql
-- Banco  : SQL Server 2014
-- Tabela : EntradaEstoque
-- Objetivo: Adicionar os campos de totais da NF-e que estavam
--           faltando no grupo <ICMSTot>:
--
-- ValorFCP      : FCP sobre ICMS normal          <vFCP>
-- ValorFCPST    : FCP retido por ST              <vFCPST>
-- ValorFCPSTRet : FCP retido anteriormente por ST <vFCPSTRet>
-- ValorIPIDevol : IPI devolvido (notas devolucao) <vIPIDevol>
--
-- Observacao: vFCP/vFCPST/vFCPSTRet existem desde NF-e 3.10
--             (NT 2015/002). vIPIDevol desde NF-e 4.00.
-- ============================================================

-- ValorFCP: Valor total do FCP sobre ICMS normal.
-- Tag XML: <vFCP> dentro de <ICMSTot>
IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('EntradaEstoque') AND name = 'ValorFCP'
)
    ALTER TABLE EntradaEstoque ADD ValorFCP DECIMAL(15,2) NOT NULL DEFAULT 0;
GO

-- ValorFCPST: Valor total do FCP retido por Substituicao Tributaria.
-- Tag XML: <vFCPST> dentro de <ICMSTot>
IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('EntradaEstoque') AND name = 'ValorFCPST'
)
    ALTER TABLE EntradaEstoque ADD ValorFCPST DECIMAL(15,2) NOT NULL DEFAULT 0;
GO

-- ValorFCPSTRet: Valor total do FCP retido anteriormente por ST.
-- Tag XML: <vFCPSTRet> dentro de <ICMSTot>
IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('EntradaEstoque') AND name = 'ValorFCPSTRet'
)
    ALTER TABLE EntradaEstoque ADD ValorFCPSTRet DECIMAL(15,2) NOT NULL DEFAULT 0;
GO

-- ValorIPIDevol: Valor total do IPI devolvido (usado em notas de devolucao).
-- Tag XML: <vIPIDevol> dentro de <ICMSTot>
IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('EntradaEstoque') AND name = 'ValorIPIDevol'
)
    ALTER TABLE EntradaEstoque ADD ValorIPIDevol DECIMAL(15,2) NOT NULL DEFAULT 0;
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
WHERE object_id = OBJECT_ID('EntradaEstoque')
  AND name IN ('ValorFCP', 'ValorFCPST', 'ValorFCPSTRet', 'ValorIPIDevol')
ORDER BY name;
GO

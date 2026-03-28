-- ============================================================
-- Script: AlterTable_EntradaEstoque_PagvPag.sql
-- Banco  : SQL Server 2014
-- Tabela : EntradaEstoque
-- Objetivo: Adicionar o campo do valor de pagamento declarado
--           pelo emitente no grupo <pag> da NF-e 4.0+:
--
-- PagvPag : Valor do pagamento declarado  <pag/detPag/vPag>
--
-- Observacao: pode diferir do valor da duplicata em casos de
-- troco ou pagamento parcial. Usa o primeiro <detPag>.
-- ============================================================

IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('EntradaEstoque') AND name = 'PagvPag'
)
    ALTER TABLE EntradaEstoque ADD PagvPag DECIMAL(15,2) NOT NULL DEFAULT 0;
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
  AND name = 'PagvPag';
GO

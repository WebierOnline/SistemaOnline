-- ============================================================
-- Script: AlterTable_EntradaEstoque_modFrete.sql
-- Banco  : SQL Server 2014
-- Tabela : EntradaEstoque
-- Objetivo: Adicionar o campo de modalidade do frete da NF-e:
--
-- TranspModFrete : Modalidade do frete  <transp/modFrete>
--
-- Valores:
--   0 = Contratacao do Frete por conta do Emitente (CIF)
--   1 = Contratacao do Frete por conta do Destinatario (FOB)
--   2 = Contratacao do Frete por conta de Terceiros
--   3 = Transporte Proprio por conta do Emitente
--   4 = Transporte Proprio por conta do Destinatario
--   9 = Sem Ocorrencia de Transporte
-- ============================================================

IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('EntradaEstoque') AND name = 'TranspModFrete'
)
    ALTER TABLE EntradaEstoque ADD TranspModFrete TINYINT NOT NULL DEFAULT 9;
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
  AND name = 'TranspModFrete';
GO

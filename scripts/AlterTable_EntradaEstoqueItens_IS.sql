-- ============================================================
-- Script: AlterTable_EntradaEstoqueItens_IS.sql
-- Banco  : SQL Server 2014
-- Tabela : EntradaEstoqueItens
-- Objetivo: Adicionar os campos do Imposto Seletivo (IS)
--           por item da NF-e (NT 2025.002 - Reforma Tributaria
--           LC 214/2025 - Grupo UB01-UB11):
--
-- ISCST       : CST do Imposto Seletivo          <CSTIS>
-- IScClassTrib: Classificacao tributaria IS      <cClassTribIS>
-- ISvBCIS     : Base de calculo do IS            <vBCIS>
-- ISpIS       : Aliquota do IS (percentual)      <pIS>
-- ISvIS       : Valor do Imposto Seletivo        <vIS>
-- ============================================================

-- ISCST: Codigo de Situacao Tributaria do Imposto Seletivo.
-- Formato: 3 digitos (ex: 001, 999)
-- Tag XML: <CSTIS> dentro de <IS> dentro de <imposto>
IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('EntradaEstoqueItens') AND name = 'ISCST'
)
    ALTER TABLE EntradaEstoqueItens ADD ISCST VARCHAR(3) NOT NULL DEFAULT '';
GO

-- IScClassTrib: Classificacao tributaria do IS.
-- Formato: 6 digitos
-- Tag XML: <cClassTribIS> dentro de <IS>
IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('EntradaEstoqueItens') AND name = 'IScClassTrib'
)
    ALTER TABLE EntradaEstoqueItens ADD IScClassTrib VARCHAR(6) NOT NULL DEFAULT '';
GO

-- ISvBCIS: Base de calculo do Imposto Seletivo.
-- Tag XML: <vBCIS> dentro de <IS> (sequencia XML opcional)
IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('EntradaEstoqueItens') AND name = 'ISvBCIS'
)
    ALTER TABLE EntradaEstoqueItens ADD ISvBCIS DECIMAL(15,2) NOT NULL DEFAULT 0;
GO

-- ISpIS: Aliquota do Imposto Seletivo em percentual.
-- Tag XML: <pIS> dentro de <IS>
IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('EntradaEstoqueItens') AND name = 'ISpIS'
)
    ALTER TABLE EntradaEstoqueItens ADD ISpIS DECIMAL(15,4) NOT NULL DEFAULT 0;
GO

-- ISvIS: Valor do Imposto Seletivo calculado.
-- Tag XML: <vIS> dentro de <IS>
IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('EntradaEstoqueItens') AND name = 'ISvIS'
)
    ALTER TABLE EntradaEstoqueItens ADD ISvIS DECIMAL(15,2) NOT NULL DEFAULT 0;
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
  AND name IN ('ISCST', 'IScClassTrib', 'ISvBCIS', 'ISpIS', 'ISvIS')
ORDER BY name;
GO

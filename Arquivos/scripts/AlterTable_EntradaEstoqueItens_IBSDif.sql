-- ============================================================
-- Script: AlterTable_EntradaEstoqueItens_IBSDif.sql
-- Banco  : SQL Server 2014
-- Tabela : EntradaEstoqueItens
-- Objetivo: Adicionar os campos de Diferimento (gDif) e
--           Devolucao de Tributos (gDevTrib) do IBS e CBS
--           por item da NF-e (NT 2025.002 - UB21-25, UB40-44,
--           UB59-63):
--
-- gIBSCBS/gIBSUF/gDif    : pDif, vDif (diferimento IBS UF)
-- gIBSCBS/gIBSUF/gDevTrib: vDevTrib (devolucao tributos UF)
-- gIBSCBS/gIBSMun/gDif   : pDif, vDif (diferimento IBS Mun)
-- gIBSCBS/gIBSMun/gDevTrib: vDevTrib (devolucao tributos Mun)
-- gIBSCBS/gCBS/gDif      : pDif, vDif (diferimento CBS)
-- gIBSCBS/gCBS/gDevTrib  : vDevTrib (devolucao tributos CBS)
-- ============================================================

-- ============================================================
-- IBS UF - Diferimento e Devolucao de Tributos
-- ============================================================

-- IBSUFpDif: Percentual do diferimento do IBS da UF.
-- Tag XML: <pDif> dentro de <gIBSCBS/gIBSUF/gDif>
IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('EntradaEstoqueItens') AND name = 'IBSUFpDif'
)
    ALTER TABLE EntradaEstoqueItens ADD IBSUFpDif DECIMAL(15,4) NOT NULL DEFAULT 0;
GO

-- IBSUFvDif: Valor do diferimento do IBS da UF.
-- Tag XML: <vDif> dentro de <gIBSCBS/gIBSUF/gDif>
IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('EntradaEstoqueItens') AND name = 'IBSUFvDif'
)
    ALTER TABLE EntradaEstoqueItens ADD IBSUFvDif DECIMAL(15,2) NOT NULL DEFAULT 0;
GO

-- IBSUFvDevTrib: Valor da devolucao de tributos do IBS UF (cashback).
-- Tag XML: <vDevTrib> dentro de <gIBSCBS/gIBSUF/gDevTrib>
IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('EntradaEstoqueItens') AND name = 'IBSUFvDevTrib'
)
    ALTER TABLE EntradaEstoqueItens ADD IBSUFvDevTrib DECIMAL(15,2) NOT NULL DEFAULT 0;
GO

-- ============================================================
-- IBS Municipal - Diferimento e Devolucao de Tributos
-- ============================================================

-- IBSMunpDif: Percentual do diferimento do IBS Municipal.
-- Tag XML: <pDif> dentro de <gIBSCBS/gIBSMun/gDif>
IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('EntradaEstoqueItens') AND name = 'IBSMunpDif'
)
    ALTER TABLE EntradaEstoqueItens ADD IBSMunpDif DECIMAL(15,4) NOT NULL DEFAULT 0;
GO

-- IBSMunvDif: Valor do diferimento do IBS Municipal.
-- Tag XML: <vDif> dentro de <gIBSCBS/gIBSMun/gDif>
IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('EntradaEstoqueItens') AND name = 'IBSMunvDif'
)
    ALTER TABLE EntradaEstoqueItens ADD IBSMunvDif DECIMAL(15,2) NOT NULL DEFAULT 0;
GO

-- IBSMunvDevTrib: Valor da devolucao de tributos do IBS Municipal.
-- Tag XML: <vDevTrib> dentro de <gIBSCBS/gIBSMun/gDevTrib>
IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('EntradaEstoqueItens') AND name = 'IBSMunvDevTrib'
)
    ALTER TABLE EntradaEstoqueItens ADD IBSMunvDevTrib DECIMAL(15,2) NOT NULL DEFAULT 0;
GO

-- ============================================================
-- CBS - Diferimento e Devolucao de Tributos
-- ============================================================

-- CBSpDif: Percentual do diferimento da CBS.
-- Tag XML: <pDif> dentro de <gIBSCBS/gCBS/gDif>
IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('EntradaEstoqueItens') AND name = 'CBSpDif'
)
    ALTER TABLE EntradaEstoqueItens ADD CBSpDif DECIMAL(15,4) NOT NULL DEFAULT 0;
GO

-- CBSvDif: Valor do diferimento da CBS.
-- Tag XML: <vDif> dentro de <gIBSCBS/gCBS/gDif>
IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('EntradaEstoqueItens') AND name = 'CBSvDif'
)
    ALTER TABLE EntradaEstoqueItens ADD CBSvDif DECIMAL(15,2) NOT NULL DEFAULT 0;
GO

-- CBSvDevTrib: Valor da devolucao de tributos da CBS.
-- Tag XML: <vDevTrib> dentro de <gIBSCBS/gCBS/gDevTrib>
IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('EntradaEstoqueItens') AND name = 'CBSvDevTrib'
)
    ALTER TABLE EntradaEstoqueItens ADD CBSvDevTrib DECIMAL(15,2) NOT NULL DEFAULT 0;
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
  AND name IN (
      'IBSUFpDif', 'IBSUFvDif', 'IBSUFvDevTrib',
      'IBSMunpDif', 'IBSMunvDif', 'IBSMunvDevTrib',
      'CBSpDif', 'CBSvDif', 'CBSvDevTrib'
  )
ORDER BY name;
GO

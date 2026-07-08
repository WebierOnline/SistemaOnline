-- ============================================================
-- Script: AlterTable_EntradaEstoqueItens_IBSCBSMono.sql
-- Banco  : SQL Server 2014
-- Tabela : EntradaEstoqueItens
-- Objetivo: Adicionar os campos do grupo gIBSCBSMono da NF-e
--           (NT 2025.002 - UB84-105 - tributacao monofasica
--           IBS/CBS, aplicavel a combustiveis e monofasicos):
--
-- gMonoPadrao : qBCMono, adRemIBS, adRemCBS, vIBSMono, vCBSMono
-- gMonoReten  : qBCMonoReten, adRemIBSReten, vIBSMonoReten,
--               adRemCBSReten, vCBSMonoReten
-- gMonoRet    : qBCMonoRet, adRemIBSRet, vIBSMonoRet,
--               adRemCBSRet, vCBSMonoRet
-- Totais item : vTotIBSMonoItem, vTotCBSMonoItem
-- ============================================================

-- ============================================================
-- gMonoPadrao - Tributacao Monofasica Padrao (UB84a-UB89)
-- ============================================================

IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('EntradaEstoqueItens') AND name = 'IBSCBSMonoqBCMono'
)
    ALTER TABLE EntradaEstoqueItens ADD IBSCBSMonoqBCMono DECIMAL(15,4) NOT NULL DEFAULT 0;
GO

IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('EntradaEstoqueItens') AND name = 'IBSCBSMonoadRemIBS'
)
    ALTER TABLE EntradaEstoqueItens ADD IBSCBSMonoadRemIBS DECIMAL(15,4) NOT NULL DEFAULT 0;
GO

IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('EntradaEstoqueItens') AND name = 'IBSCBSMonoadRemCBS'
)
    ALTER TABLE EntradaEstoqueItens ADD IBSCBSMonoadRemCBS DECIMAL(15,4) NOT NULL DEFAULT 0;
GO

IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('EntradaEstoqueItens') AND name = 'IBSCBSMonovIBSMono'
)
    ALTER TABLE EntradaEstoqueItens ADD IBSCBSMonovIBSMono DECIMAL(15,2) NOT NULL DEFAULT 0;
GO

IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('EntradaEstoqueItens') AND name = 'IBSCBSMonovCBSMono'
)
    ALTER TABLE EntradaEstoqueItens ADD IBSCBSMonovCBSMono DECIMAL(15,2) NOT NULL DEFAULT 0;
GO

-- ============================================================
-- gMonoReten - Tributacao Monofasica Sujeita a Retencao (UB90-93b)
-- ============================================================

IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('EntradaEstoqueItens') AND name = 'IBSCBSMonoqBCMonoReten'
)
    ALTER TABLE EntradaEstoqueItens ADD IBSCBSMonoqBCMonoReten DECIMAL(15,4) NOT NULL DEFAULT 0;
GO

IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('EntradaEstoqueItens') AND name = 'IBSCBSMonoadRemIBSReten'
)
    ALTER TABLE EntradaEstoqueItens ADD IBSCBSMonoadRemIBSReten DECIMAL(15,4) NOT NULL DEFAULT 0;
GO

IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('EntradaEstoqueItens') AND name = 'IBSCBSMonovIBSMonoReten'
)
    ALTER TABLE EntradaEstoqueItens ADD IBSCBSMonovIBSMonoReten DECIMAL(15,2) NOT NULL DEFAULT 0;
GO

IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('EntradaEstoqueItens') AND name = 'IBSCBSMonoadRemCBSReten'
)
    ALTER TABLE EntradaEstoqueItens ADD IBSCBSMonoadRemCBSReten DECIMAL(15,4) NOT NULL DEFAULT 0;
GO

IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('EntradaEstoqueItens') AND name = 'IBSCBSMonovCBSMonoReten'
)
    ALTER TABLE EntradaEstoqueItens ADD IBSCBSMonovCBSMonoReten DECIMAL(15,2) NOT NULL DEFAULT 0;
GO

-- ============================================================
-- gMonoRet - Tributacao Monofasica Retida Anteriormente (UB94-98a)
-- ============================================================

IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('EntradaEstoqueItens') AND name = 'IBSCBSMonoqBCMonoRet'
)
    ALTER TABLE EntradaEstoqueItens ADD IBSCBSMonoqBCMonoRet DECIMAL(15,4) NOT NULL DEFAULT 0;
GO

IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('EntradaEstoqueItens') AND name = 'IBSCBSMonoadRemIBSRet'
)
    ALTER TABLE EntradaEstoqueItens ADD IBSCBSMonoadRemIBSRet DECIMAL(15,4) NOT NULL DEFAULT 0;
GO

IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('EntradaEstoqueItens') AND name = 'IBSCBSMonovIBSMonoRet'
)
    ALTER TABLE EntradaEstoqueItens ADD IBSCBSMonovIBSMonoRet DECIMAL(15,2) NOT NULL DEFAULT 0;
GO

IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('EntradaEstoqueItens') AND name = 'IBSCBSMonoadRemCBSRet'
)
    ALTER TABLE EntradaEstoqueItens ADD IBSCBSMonoadRemCBSRet DECIMAL(15,4) NOT NULL DEFAULT 0;
GO

IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('EntradaEstoqueItens') AND name = 'IBSCBSMonovCBSMonoRet'
)
    ALTER TABLE EntradaEstoqueItens ADD IBSCBSMonovCBSMonoRet DECIMAL(15,2) NOT NULL DEFAULT 0;
GO

-- ============================================================
-- Totais monofasicos do item (UB104-UB105)
-- ============================================================

-- IBSCBSMonovTotIBSMonoItem: Total de IBS Monofasico do item.
-- Tag XML: <vTotIBSMonoItem> dentro de <IBSCBS/gIBSCBSMono>
IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('EntradaEstoqueItens') AND name = 'IBSCBSMonovTotIBSMonoItem'
)
    ALTER TABLE EntradaEstoqueItens ADD IBSCBSMonovTotIBSMonoItem DECIMAL(15,2) NOT NULL DEFAULT 0;
GO

-- IBSCBSMonovTotCBSMonoItem: Total de CBS Monofasica do item.
-- Tag XML: <vTotCBSMonoItem> dentro de <IBSCBS/gIBSCBSMono>
IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('EntradaEstoqueItens') AND name = 'IBSCBSMonovTotCBSMonoItem'
)
    ALTER TABLE EntradaEstoqueItens ADD IBSCBSMonovTotCBSMonoItem DECIMAL(15,2) NOT NULL DEFAULT 0;
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
      'IBSCBSMonoqBCMono', 'IBSCBSMonoadRemIBS', 'IBSCBSMonoadRemCBS',
      'IBSCBSMonovIBSMono', 'IBSCBSMonovCBSMono',
      'IBSCBSMonoqBCMonoReten', 'IBSCBSMonoadRemIBSReten', 'IBSCBSMonovIBSMonoReten',
      'IBSCBSMonoadRemCBSReten', 'IBSCBSMonovCBSMonoReten',
      'IBSCBSMonoqBCMonoRet', 'IBSCBSMonoadRemIBSRet', 'IBSCBSMonovIBSMonoRet',
      'IBSCBSMonoadRemCBSRet', 'IBSCBSMonovCBSMonoRet',
      'IBSCBSMonovTotIBSMonoItem', 'IBSCBSMonovTotCBSMonoItem'
  )
ORDER BY name;
GO

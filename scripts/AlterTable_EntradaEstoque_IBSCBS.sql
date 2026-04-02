-- ============================================================
-- Script: AlterTable_EntradaEstoque_IBSCBS.sql
-- Banco  : SQL Server 2014
-- Tabela : EntradaEstoque
-- Objetivo: Adicionar os campos de totais do grupo <IBSCBSTot>
--           da NF-e (Reforma Tributaria - EC 132/2023 + LC 214/2025):
--           totais IBS UF, IBS Mun, IBS geral e CBS + vNFTot
-- ============================================================

-- ============================================================
-- Base de calculo total IBS+CBS
-- ============================================================

-- TotIBSCBSvBC: Base de calculo total IBS+CBS.
-- Tag XML: <vBC> dentro de <total/IBSCBSTot>
IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('EntradaEstoque') AND name = 'TotIBSCBSvBC'
)
    ALTER TABLE EntradaEstoque ADD TotIBSCBSvBC DECIMAL(15,2) NOT NULL DEFAULT 0;
GO

-- ============================================================
-- Totais IBS UF (gIBSTot/gIBSUF)
-- ============================================================

-- TotIBSUFvDif: Valor do IBS UF diferido.
-- Tag XML: <vDif> dentro de <IBSCBSTot/gIBSTot/gIBSUF>
IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('EntradaEstoque') AND name = 'TotIBSUFvDif'
)
    ALTER TABLE EntradaEstoque ADD TotIBSUFvDif DECIMAL(15,2) NOT NULL DEFAULT 0;
GO

-- TotIBSUFvDevTrib: Valor do IBS UF a devolver ao contribuinte.
-- Tag XML: <vDevTrib> dentro de <IBSCBSTot/gIBSTot/gIBSUF>
IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('EntradaEstoque') AND name = 'TotIBSUFvDevTrib'
)
    ALTER TABLE EntradaEstoque ADD TotIBSUFvDevTrib DECIMAL(15,2) NOT NULL DEFAULT 0;
GO

-- TotIBSUFvIBS: Valor total do IBS estadual.
-- Tag XML: <vIBS> dentro de <IBSCBSTot/gIBSTot/gIBSUF>
IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('EntradaEstoque') AND name = 'TotIBSUFvIBS'
)
    ALTER TABLE EntradaEstoque ADD TotIBSUFvIBS DECIMAL(15,2) NOT NULL DEFAULT 0;
GO

-- ============================================================
-- Totais IBS Mun (gIBSTot/gIBSMun)
-- ============================================================

-- TotIBSMunvDif: Valor do IBS municipal diferido.
-- Tag XML: <vDif> dentro de <IBSCBSTot/gIBSTot/gIBSMun>
IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('EntradaEstoque') AND name = 'TotIBSMunvDif'
)
    ALTER TABLE EntradaEstoque ADD TotIBSMunvDif DECIMAL(15,2) NOT NULL DEFAULT 0;
GO

-- TotIBSMunvDevTrib: Valor do IBS municipal a devolver ao contribuinte.
-- Tag XML: <vDevTrib> dentro de <IBSCBSTot/gIBSTot/gIBSMun>
IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('EntradaEstoque') AND name = 'TotIBSMunvDevTrib'
)
    ALTER TABLE EntradaEstoque ADD TotIBSMunvDevTrib DECIMAL(15,2) NOT NULL DEFAULT 0;
GO

-- TotIBSMunvIBS: Valor total do IBS municipal.
-- Tag XML: <vIBS> dentro de <IBSCBSTot/gIBSTot/gIBSMun>
IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('EntradaEstoque') AND name = 'TotIBSMunvIBS'
)
    ALTER TABLE EntradaEstoque ADD TotIBSMunvIBS DECIMAL(15,2) NOT NULL DEFAULT 0;
GO

-- ============================================================
-- Total IBS geral (gIBSTot)
-- ============================================================

-- TotIBSvIBS: Valor total do IBS (UF + Municipal).
-- Tag XML: <vIBS> dentro de <IBSCBSTot/gIBSTot>
IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('EntradaEstoque') AND name = 'TotIBSvIBS'
)
    ALTER TABLE EntradaEstoque ADD TotIBSvIBS DECIMAL(15,2) NOT NULL DEFAULT 0;
GO

-- TotIBSvCredPres: Valor do credito presumido de IBS.
-- Tag XML: <vCredPres> dentro de <IBSCBSTot/gIBSTot>
IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('EntradaEstoque') AND name = 'TotIBSvCredPres'
)
    ALTER TABLE EntradaEstoque ADD TotIBSvCredPres DECIMAL(15,2) NOT NULL DEFAULT 0;
GO

-- TotIBSvCredPresCondSus: Valor do credito presumido de IBS com condicao suspensiva.
-- Tag XML: <vCredPresCondSus> dentro de <IBSCBSTot/gIBSTot>
IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('EntradaEstoque') AND name = 'TotIBSvCredPresCondSus'
)
    ALTER TABLE EntradaEstoque ADD TotIBSvCredPresCondSus DECIMAL(15,2) NOT NULL DEFAULT 0;
GO

-- ============================================================
-- Totais CBS (gCBSTot)
-- ============================================================

-- TotCBSvDif: Valor do CBS diferido.
-- Tag XML: <vDif> dentro de <IBSCBSTot/gCBSTot>
IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('EntradaEstoque') AND name = 'TotCBSvDif'
)
    ALTER TABLE EntradaEstoque ADD TotCBSvDif DECIMAL(15,2) NOT NULL DEFAULT 0;
GO

-- TotCBSvDevTrib: Valor do CBS a devolver ao contribuinte.
-- Tag XML: <vDevTrib> dentro de <IBSCBSTot/gCBSTot>
IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('EntradaEstoque') AND name = 'TotCBSvDevTrib'
)
    ALTER TABLE EntradaEstoque ADD TotCBSvDevTrib DECIMAL(15,2) NOT NULL DEFAULT 0;
GO

-- TotCBSvCBS: Valor total do CBS.
-- Tag XML: <vCBS> dentro de <IBSCBSTot/gCBSTot>
IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('EntradaEstoque') AND name = 'TotCBSvCBS'
)
    ALTER TABLE EntradaEstoque ADD TotCBSvCBS DECIMAL(15,2) NOT NULL DEFAULT 0;
GO

-- TotCBSvCredPres: Valor do credito presumido de CBS.
-- Tag XML: <vCredPres> dentro de <IBSCBSTot/gCBSTot>
IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('EntradaEstoque') AND name = 'TotCBSvCredPres'
)
    ALTER TABLE EntradaEstoque ADD TotCBSvCredPres DECIMAL(15,2) NOT NULL DEFAULT 0;
GO

-- TotCBSvCredPresCondSus: Valor do credito presumido de CBS com condicao suspensiva.
-- Tag XML: <vCredPresCondSus> dentro de <IBSCBSTot/gCBSTot>
IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('EntradaEstoque') AND name = 'TotCBSvCredPresCondSus'
)
    ALTER TABLE EntradaEstoque ADD TotCBSvCredPresCondSus DECIMAL(15,2) NOT NULL DEFAULT 0;
GO

-- ============================================================
-- Valor total da NF-e incluindo IBS+CBS
-- ============================================================

-- vNFTot: Valor total da NF-e incluindo IBS+CBS (pode diferir de ValorNota/vNF para NF-e com reforma tributaria).
-- Tag XML: <vNFTot> dentro de <total/IBSCBSTot>
IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('EntradaEstoque') AND name = 'vNFTot'
)
    ALTER TABLE EntradaEstoque ADD vNFTot DECIMAL(15,2) NOT NULL DEFAULT 0;
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
  AND name IN (
      'TotIBSCBSvBC',
      'TotIBSUFvDif', 'TotIBSUFvDevTrib', 'TotIBSUFvIBS',
      'TotIBSMunvDif', 'TotIBSMunvDevTrib', 'TotIBSMunvIBS',
      'TotIBSvIBS', 'TotIBSvCredPres', 'TotIBSvCredPresCondSus',
      'TotCBSvDif', 'TotCBSvDevTrib', 'TotCBSvCBS',
      'TotCBSvCredPres', 'TotCBSvCredPresCondSus',
      'vNFTot'
  )
ORDER BY name;
GO

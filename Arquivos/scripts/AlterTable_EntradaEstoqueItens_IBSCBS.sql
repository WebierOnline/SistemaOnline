-- ============================================================
-- Script: AlterTable_EntradaEstoqueItens_IBSCBS.sql
-- Banco  : SQL Server 2014
-- Tabela : EntradaEstoqueItens
-- Objetivo: Adicionar os campos do grupo <IBSCBS> da NF-e
--           (Reforma Tributaria - EC 132/2023 + LC 214/2025):
--           IBS (UF + Municipal) e CBS por item
-- ============================================================

-- IBSCBSCST: Codigo de Situacao Tributaria IBS/CBS.
-- Formato: 3 digitos (ex: 000, 200)
-- Tag XML: <CST> dentro de <IBSCBS>
IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('EntradaEstoqueItens') AND name = 'IBSCBSCST'
)
    ALTER TABLE EntradaEstoqueItens ADD IBSCBSCST VARCHAR(3) NOT NULL DEFAULT '';
GO

-- IBSCBScClassTrib: Classificacao tributaria do produto (6 digitos).
-- Tag XML: <cClassTrib> dentro de <IBSCBS>
IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('EntradaEstoqueItens') AND name = 'IBSCBScClassTrib'
)
    ALTER TABLE EntradaEstoqueItens ADD IBSCBScClassTrib VARCHAR(6) NOT NULL DEFAULT '';
GO

-- IBSCBSvBC: Base de calculo IBS+CBS.
-- Tag XML: <vBC> dentro de <IBSCBS/gIBSCBS>
IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('EntradaEstoqueItens') AND name = 'IBSCBSvBC'
)
    ALTER TABLE EntradaEstoqueItens ADD IBSCBSvBC DECIMAL(15,2) NOT NULL DEFAULT 0;
GO

-- ============================================================
-- IBS UF (gIBSCBS/gIBSUF)
-- ============================================================

-- IBSUFpAliq: Aliquota IBS estadual (%).
-- Tag XML: <pIBSUF> dentro de <gIBSCBS/gIBSUF>
IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('EntradaEstoqueItens') AND name = 'IBSUFpAliq'
)
    ALTER TABLE EntradaEstoqueItens ADD IBSUFpAliq DECIMAL(7,4) NOT NULL DEFAULT 0;
GO

-- IBSUFpRedAliq: Percentual de reducao da aliquota IBS UF (opcional - CST com reducao).
-- Tag XML: <pRedAliq> dentro de <gIBSCBS/gIBSUF/gRed>
IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('EntradaEstoqueItens') AND name = 'IBSUFpRedAliq'
)
    ALTER TABLE EntradaEstoqueItens ADD IBSUFpRedAliq DECIMAL(7,4) NOT NULL DEFAULT 0;
GO

-- IBSUFpAliqEfet: Aliquota efetiva IBS UF apos aplicar reducao (opcional).
-- Tag XML: <pAliqEfet> dentro de <gIBSCBS/gIBSUF/gRed>
IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('EntradaEstoqueItens') AND name = 'IBSUFpAliqEfet'
)
    ALTER TABLE EntradaEstoqueItens ADD IBSUFpAliqEfet DECIMAL(7,4) NOT NULL DEFAULT 0;
GO

-- IBSUFvIBS: Valor do IBS estadual.
-- Tag XML: <vIBSUF> dentro de <gIBSCBS/gIBSUF>
IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('EntradaEstoqueItens') AND name = 'IBSUFvIBS'
)
    ALTER TABLE EntradaEstoqueItens ADD IBSUFvIBS DECIMAL(15,2) NOT NULL DEFAULT 0;
GO

-- ============================================================
-- IBS Mun (gIBSCBS/gIBSMun)
-- ============================================================

-- IBSMunpAliq: Aliquota IBS municipal (%).
-- Tag XML: <pIBSMun> dentro de <gIBSCBS/gIBSMun>
IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('EntradaEstoqueItens') AND name = 'IBSMunpAliq'
)
    ALTER TABLE EntradaEstoqueItens ADD IBSMunpAliq DECIMAL(7,4) NOT NULL DEFAULT 0;
GO

-- IBSMunpRedAliq: Percentual de reducao da aliquota IBS municipal (opcional).
-- Tag XML: <pRedAliq> dentro de <gIBSCBS/gIBSMun/gRed>
IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('EntradaEstoqueItens') AND name = 'IBSMunpRedAliq'
)
    ALTER TABLE EntradaEstoqueItens ADD IBSMunpRedAliq DECIMAL(7,4) NOT NULL DEFAULT 0;
GO

-- IBSMunpAliqEfet: Aliquota efetiva IBS municipal apos reducao (opcional).
-- Tag XML: <pAliqEfet> dentro de <gIBSCBS/gIBSMun/gRed>
IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('EntradaEstoqueItens') AND name = 'IBSMunpAliqEfet'
)
    ALTER TABLE EntradaEstoqueItens ADD IBSMunpAliqEfet DECIMAL(7,4) NOT NULL DEFAULT 0;
GO

-- IBSMunvIBS: Valor do IBS municipal.
-- Tag XML: <vIBSMun> dentro de <gIBSCBS/gIBSMun>
IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('EntradaEstoqueItens') AND name = 'IBSMunvIBS'
)
    ALTER TABLE EntradaEstoqueItens ADD IBSMunvIBS DECIMAL(15,2) NOT NULL DEFAULT 0;
GO

-- IBSvIBS: Valor total do IBS (UF + Municipal).
-- Tag XML: <vIBS> dentro de <gIBSCBS>
IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('EntradaEstoqueItens') AND name = 'IBSvIBS'
)
    ALTER TABLE EntradaEstoqueItens ADD IBSvIBS DECIMAL(15,2) NOT NULL DEFAULT 0;
GO

-- ============================================================
-- CBS (gIBSCBS/gCBS)
-- ============================================================

-- CBSpAliq: Aliquota CBS (%).
-- Tag XML: <pCBS> dentro de <gIBSCBS/gCBS>
IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('EntradaEstoqueItens') AND name = 'CBSpAliq'
)
    ALTER TABLE EntradaEstoqueItens ADD CBSpAliq DECIMAL(7,4) NOT NULL DEFAULT 0;
GO

-- CBSpRedAliq: Percentual de reducao da aliquota CBS (opcional).
-- Tag XML: <pRedAliq> dentro de <gIBSCBS/gCBS/gRed>
IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('EntradaEstoqueItens') AND name = 'CBSpRedAliq'
)
    ALTER TABLE EntradaEstoqueItens ADD CBSpRedAliq DECIMAL(7,4) NOT NULL DEFAULT 0;
GO

-- CBSpAliqEfet: Aliquota efetiva CBS apos aplicar reducao (opcional).
-- Tag XML: <pAliqEfet> dentro de <gIBSCBS/gCBS/gRed>
IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('EntradaEstoqueItens') AND name = 'CBSpAliqEfet'
)
    ALTER TABLE EntradaEstoqueItens ADD CBSpAliqEfet DECIMAL(7,4) NOT NULL DEFAULT 0;
GO

-- CBSvCBS: Valor do CBS.
-- Tag XML: <vCBS> dentro de <gIBSCBS/gCBS>
IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('EntradaEstoqueItens') AND name = 'CBSvCBS'
)
    ALTER TABLE EntradaEstoqueItens ADD CBSvCBS DECIMAL(15,2) NOT NULL DEFAULT 0;
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
      'IBSCBSCST', 'IBSCBScClassTrib', 'IBSCBSvBC',
      'IBSUFpAliq', 'IBSUFpRedAliq', 'IBSUFpAliqEfet', 'IBSUFvIBS',
      'IBSMunpAliq', 'IBSMunpRedAliq', 'IBSMunpAliqEfet', 'IBSMunvIBS',
      'IBSvIBS',
      'CBSpAliq', 'CBSpRedAliq', 'CBSpAliqEfet', 'CBSvCBS'
  )
ORDER BY name;
GO

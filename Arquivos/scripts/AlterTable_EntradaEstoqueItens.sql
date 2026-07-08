-- ============================================================
-- Script: AlterTable_EntradaEstoqueItens.sql
-- Banco  : SQL Server 2014
-- Tabela : EntradaEstoqueItens
-- Objetivo: Adicionar os campos de impostos que existem na
--           NF-e (XML) mas nao possuiam coluna na tabela.
-- Campos adicionados por grupo fiscal:
--   ICMS51 (Diferimento): vICMSOp, pDif
--   ICMS40/41 (Desoneracao): vICMSDeson, motDesICMS
--   FCP-ST (ICMS10/70/90): vBCFCPST, pFCPST, vFCPST
--   ICMS60 (ST Retido): vBCSTRet, pST, vICMSSubstituto,
--                        vICMSSTRet, vBCFCPSTRet, pFCPSTRet, vFCPSTRet
-- ============================================================

-- ============================================================
-- ICMS51 - Diferimento
-- ============================================================

-- vICMSOp: Valor do ICMS da operacao antes de aplicar o diferimento.
-- Tag XML: <vICMSOp> dentro de <ICMS51>
IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('EntradaEstoqueItens') AND name = 'vICMSOp'
)
    ALTER TABLE EntradaEstoqueItens ADD vICMSOp DECIMAL(15,2) NOT NULL DEFAULT 0;
GO

-- pDif: Percentual do diferimento do ICMS (ex: 100.00 = diferimento total).
-- Tag XML: <pDif> dentro de <ICMS51>
IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('EntradaEstoqueItens') AND name = 'pDif'
)
    ALTER TABLE EntradaEstoqueItens ADD pDif DECIMAL(5,2) NOT NULL DEFAULT 0;
GO

-- ============================================================
-- ICMS40 / ICMS41 - Desoneracao do ICMS
-- ============================================================

-- vICMSDeson: Valor do ICMS desonerado (isento, nao tributado ou com reducao).
-- Tag XML: <vICMSDeson> dentro de <ICMS40> ou <ICMS41>
IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('EntradaEstoqueItens') AND name = 'vICMSDeson'
)
    ALTER TABLE EntradaEstoqueItens ADD vICMSDeson DECIMAL(15,2) NOT NULL DEFAULT 0;
GO

-- motDesICMS: Motivo da desoneracao do ICMS.
-- Valores: 01=Taxi, 02=Deficiente Fisico, 03=Produtor Agropecuario,
--          04=Frotista/Locadora, 05=Diplomatico/Consular,
--          06=Utilitarios e Motocicletas Amazonia, 07=SUFRAMA,
--          08=Venda a orgao publico, 09=Outros, 10=Deficiente Condutor,
--          11=Deficiente nao condutor, 12=Orgao de Fomento Agropecuario
-- Tag XML: <motDesICMS> dentro de <ICMS40> ou <ICMS41>
IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('EntradaEstoqueItens') AND name = 'motDesICMS'
)
    ALTER TABLE EntradaEstoqueItens ADD motDesICMS VARCHAR(2) NOT NULL DEFAULT '';
GO

-- ============================================================
-- FCP-ST - Fundo de Combate a Pobreza retido por Substituicao Tributaria
-- Presente nos subgrupos ICMS10, ICMS70 e ICMS90
-- ============================================================

-- vBCFCPST: Base de calculo do FCP retido por substituicao tributaria.
-- Tag XML: <vBCFCPST> dentro de <ICMS10>, <ICMS70> ou <ICMS90>
IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('EntradaEstoqueItens') AND name = 'vBCFCPST'
)
    ALTER TABLE EntradaEstoqueItens ADD vBCFCPST DECIMAL(15,2) NOT NULL DEFAULT 0;
GO

-- pFCPST: Percentual do FCP retido por substituicao tributaria.
-- Tag XML: <pFCPST> dentro de <ICMS10>, <ICMS70> ou <ICMS90>
IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('EntradaEstoqueItens') AND name = 'pFCPST'
)
    ALTER TABLE EntradaEstoqueItens ADD pFCPST DECIMAL(5,2) NOT NULL DEFAULT 0;
GO

-- vFCPST: Valor do FCP retido por substituicao tributaria.
-- Tag XML: <vFCPST> dentro de <ICMS10>, <ICMS70> ou <ICMS90>
IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('EntradaEstoqueItens') AND name = 'vFCPST'
)
    ALTER TABLE EntradaEstoqueItens ADD vFCPST DECIMAL(15,2) NOT NULL DEFAULT 0;
GO

-- ============================================================
-- ICMS60 - ICMS cobrado por ST retido anteriormente
-- ============================================================

-- vBCSTRet: Base de calculo do ICMS ST que foi retido anteriormente.
-- Tag XML: <vBCSTRet> dentro de <ICMS60>
IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('EntradaEstoqueItens') AND name = 'vBCSTRet'
)
    ALTER TABLE EntradaEstoqueItens ADD vBCSTRet DECIMAL(15,2) NOT NULL DEFAULT 0;
GO

-- pST: Aliquota suportada pelo consumidor final.
-- Tag XML: <pST> dentro de <ICMS60>
IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('EntradaEstoqueItens') AND name = 'pST'
)
    ALTER TABLE EntradaEstoqueItens ADD pST DECIMAL(5,2) NOT NULL DEFAULT 0;
GO

-- vICMSSubstituto: Valor do ICMS proprio do substituto tributario.
-- Tag XML: <vICMSSubstituto> dentro de <ICMS60>
IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('EntradaEstoqueItens') AND name = 'vICMSSubstituto'
)
    ALTER TABLE EntradaEstoqueItens ADD vICMSSubstituto DECIMAL(15,2) NOT NULL DEFAULT 0;
GO

-- vICMSSTRet: Valor do ICMS ST retido anteriormente.
-- Tag XML: <vICMSSTRet> dentro de <ICMS60>
IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('EntradaEstoqueItens') AND name = 'vICMSSTRet'
)
    ALTER TABLE EntradaEstoqueItens ADD vICMSSTRet DECIMAL(15,2) NOT NULL DEFAULT 0;
GO

-- vBCFCPSTRet: Base de calculo do FCP retido anteriormente por ST.
-- Tag XML: <vBCFCPSTRet> dentro de <ICMS60>
IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('EntradaEstoqueItens') AND name = 'vBCFCPSTRet'
)
    ALTER TABLE EntradaEstoqueItens ADD vBCFCPSTRet DECIMAL(15,2) NOT NULL DEFAULT 0;
GO

-- pFCPSTRet: Percentual do FCP retido anteriormente por ST.
-- Tag XML: <pFCPSTRet> dentro de <ICMS60>
IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('EntradaEstoqueItens') AND name = 'pFCPSTRet'
)
    ALTER TABLE EntradaEstoqueItens ADD pFCPSTRet DECIMAL(5,2) NOT NULL DEFAULT 0;
GO

-- vFCPSTRet: Valor do FCP retido anteriormente por ST.
-- Tag XML: <vFCPSTRet> dentro de <ICMS60>
IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('EntradaEstoqueItens') AND name = 'vFCPSTRet'
)
    ALTER TABLE EntradaEstoqueItens ADD vFCPSTRet DECIMAL(15,2) NOT NULL DEFAULT 0;
GO

-- ============================================================
-- Verificacao final: lista todos os campos adicionados
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
      'vICMSOp', 'pDif',
      'vICMSDeson', 'motDesICMS',
      'vBCFCPST', 'pFCPST', 'vFCPST',
      'vBCSTRet', 'pST', 'vICMSSubstituto',
      'vICMSSTRet', 'vBCFCPSTRet', 'pFCPSTRet', 'vFCPSTRet'
  )
ORDER BY name;
GO

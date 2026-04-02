-- ============================================================
-- Script: AlterTable_EntradaEstoque_IS_IBSMono.sql
-- Banco  : SQL Server 2014
-- Tabela : EntradaEstoque
-- Objetivo: Adicionar os campos de totais do Imposto Seletivo
--           (ISTot) e da tributacao monofasica IBS/CBS
--           (IBSCBSTot/gMono) da NF-e (NT 2025.002):
--
-- TotISvIS             : Total do Imposto Seletivo     <ISTot/vIS>
-- TotIBSMonovIBSMono   : Total IBS Monofasico          <gMono/vIBSMono>
-- TotIBSMonovCBSMono   : Total CBS Monofasica          <gMono/vCBSMono>
-- TotIBSMonovIBSMonoReten : Total IBS Mono c/ retencao <gMono/vIBSMonoReten>
-- TotIBSMonovCBSMonoReten : Total CBS Mono c/ retencao <gMono/vCBSMonoReten>
-- TotIBSMonovIBSMonoRet: Total IBS Mono retido ant.    <gMono/vIBSMonoRet>
-- TotIBSMonovCBSMonoRet: Total CBS Mono retida ant.    <gMono/vCBSMonoRet>
-- ============================================================

-- TotISvIS: Valor total do Imposto Seletivo da NF-e.
-- Tag XML: <vIS> dentro de <total/ISTot>
IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('EntradaEstoque') AND name = 'TotISvIS'
)
    ALTER TABLE EntradaEstoque ADD TotISvIS DECIMAL(15,2) NOT NULL DEFAULT 0;
GO

-- ============================================================
-- IBSCBSTot/gMono - Totais da tributacao monofasica IBS/CBS
-- ============================================================

-- TotIBSMonovIBSMono: Total do IBS monofasico da NF-e.
-- Tag XML: <vIBSMono> dentro de <total/IBSCBSTot/gMono>
IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('EntradaEstoque') AND name = 'TotIBSMonovIBSMono'
)
    ALTER TABLE EntradaEstoque ADD TotIBSMonovIBSMono DECIMAL(15,2) NOT NULL DEFAULT 0;
GO

-- TotIBSMonovCBSMono: Total da CBS monofasica da NF-e.
-- Tag XML: <vCBSMono> dentro de <total/IBSCBSTot/gMono>
IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('EntradaEstoque') AND name = 'TotIBSMonovCBSMono'
)
    ALTER TABLE EntradaEstoque ADD TotIBSMonovCBSMono DECIMAL(15,2) NOT NULL DEFAULT 0;
GO

-- TotIBSMonovIBSMonoReten: Total do IBS monofasico sujeito a retencao.
-- Tag XML: <vIBSMonoReten> dentro de <total/IBSCBSTot/gMono>
IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('EntradaEstoque') AND name = 'TotIBSMonovIBSMonoReten'
)
    ALTER TABLE EntradaEstoque ADD TotIBSMonovIBSMonoReten DECIMAL(15,2) NOT NULL DEFAULT 0;
GO

-- TotIBSMonovCBSMonoReten: Total da CBS monofasica sujeita a retencao.
-- Tag XML: <vCBSMonoReten> dentro de <total/IBSCBSTot/gMono>
IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('EntradaEstoque') AND name = 'TotIBSMonovCBSMonoReten'
)
    ALTER TABLE EntradaEstoque ADD TotIBSMonovCBSMonoReten DECIMAL(15,2) NOT NULL DEFAULT 0;
GO

-- TotIBSMonovIBSMonoRet: Total do IBS monofasico retido anteriormente.
-- Tag XML: <vIBSMonoRet> dentro de <total/IBSCBSTot/gMono>
IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('EntradaEstoque') AND name = 'TotIBSMonovIBSMonoRet'
)
    ALTER TABLE EntradaEstoque ADD TotIBSMonovIBSMonoRet DECIMAL(15,2) NOT NULL DEFAULT 0;
GO

-- TotIBSMonovCBSMonoRet: Total da CBS monofasica retida anteriormente.
-- Tag XML: <vCBSMonoRet> dentro de <total/IBSCBSTot/gMono>
IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('EntradaEstoque') AND name = 'TotIBSMonovCBSMonoRet'
)
    ALTER TABLE EntradaEstoque ADD TotIBSMonovCBSMonoRet DECIMAL(15,2) NOT NULL DEFAULT 0;
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
      'TotISvIS',
      'TotIBSMonovIBSMono', 'TotIBSMonovCBSMono',
      'TotIBSMonovIBSMonoReten', 'TotIBSMonovCBSMonoReten',
      'TotIBSMonovIBSMonoRet', 'TotIBSMonovCBSMonoRet'
  )
ORDER BY name;
GO

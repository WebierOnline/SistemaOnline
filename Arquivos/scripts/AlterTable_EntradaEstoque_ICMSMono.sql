-- ============================================================
-- Script: AlterTable_EntradaEstoque_ICMSMono.sql
-- Banco  : SQL Server 2014
-- Tabela : EntradaEstoque
-- Objetivo: Adicionar os campos do ICMS Monofasico e total de
--           tributos do grupo <ICMSTot> da NF-e 4.00+
--           (NT 2021/004 - combustiveis e monofasicos):
--
-- qBCMono       : Qtde BC ICMS Monofasico           <qBCMono>
-- vICMSMono     : Valor ICMS Monofasico              <vICMSMono>
-- qBCMonoReten  : Qtde BC Mono retido por ST         <qBCMonoReten>
-- vICMSMonoReten: Valor ICMS Mono retido por ST      <vICMSMonoReten>
-- qBCMonoRet    : Qtde BC Mono retido anteriormente  <qBCMonoRet>
-- vICMSMonoRet  : Valor ICMS Mono ret. anteriormente <vICMSMonoRet>
-- TotTributos   : Total aprox. tributos da NF-e      <vTotTrib>
-- ============================================================

IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('EntradaEstoque') AND name = 'qBCMono'
)
    ALTER TABLE EntradaEstoque ADD qBCMono DECIMAL(15,3) NOT NULL DEFAULT 0;
GO

IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('EntradaEstoque') AND name = 'vICMSMono'
)
    ALTER TABLE EntradaEstoque ADD vICMSMono DECIMAL(15,2) NOT NULL DEFAULT 0;
GO

IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('EntradaEstoque') AND name = 'qBCMonoReten'
)
    ALTER TABLE EntradaEstoque ADD qBCMonoReten DECIMAL(15,3) NOT NULL DEFAULT 0;
GO

IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('EntradaEstoque') AND name = 'vICMSMonoReten'
)
    ALTER TABLE EntradaEstoque ADD vICMSMonoReten DECIMAL(15,2) NOT NULL DEFAULT 0;
GO

IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('EntradaEstoque') AND name = 'qBCMonoRet'
)
    ALTER TABLE EntradaEstoque ADD qBCMonoRet DECIMAL(15,3) NOT NULL DEFAULT 0;
GO

IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('EntradaEstoque') AND name = 'vICMSMonoRet'
)
    ALTER TABLE EntradaEstoque ADD vICMSMonoRet DECIMAL(15,2) NOT NULL DEFAULT 0;
GO

-- TotTributos: Valor aproximado total dos tributos da NF-e (IBPT).
-- Tag XML: <vTotTrib> dentro de <ICMSTot>
IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('EntradaEstoque') AND name = 'TotTributos'
)
    ALTER TABLE EntradaEstoque ADD TotTributos DECIMAL(15,2) NOT NULL DEFAULT 0;
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
      'qBCMono', 'vICMSMono',
      'qBCMonoReten', 'vICMSMonoReten',
      'qBCMonoRet', 'vICMSMonoRet',
      'TotTributos'
  )
ORDER BY name;
GO

-- ============================================================
-- Script: AlterTable_EntradaEstoque_DIFAL.sql
-- Banco  : SQL Server 2014
-- Tabela : EntradaEstoque
-- Objetivo: Adicionar os campos do ICMS Diferencial de Aliquota
--           (DIFAL - EC 87/2015 / NT 2015/003) do grupo <ICMSTot>:
--
-- vFCPUFDest  : Valor total do FCP para a UF de destino     <vFCPUFDest>
-- vICMSUFDest : Valor total do ICMS para a UF de destino    <vICMSUFDest>
-- vICMSUFRemet: Valor total do ICMS para a UF de origem     <vICMSUFRemet>
--
-- Presentes em NF-e de operacoes interestaduais destinadas
-- a consumidor final nao contribuinte do ICMS (B2C).
-- ============================================================

IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('EntradaEstoque') AND name = 'vFCPUFDest'
)
    ALTER TABLE EntradaEstoque ADD vFCPUFDest DECIMAL(15,2) NOT NULL DEFAULT 0;
GO

IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('EntradaEstoque') AND name = 'vICMSUFDest'
)
    ALTER TABLE EntradaEstoque ADD vICMSUFDest DECIMAL(15,2) NOT NULL DEFAULT 0;
GO

IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('EntradaEstoque') AND name = 'vICMSUFRemet'
)
    ALTER TABLE EntradaEstoque ADD vICMSUFRemet DECIMAL(15,2) NOT NULL DEFAULT 0;
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
  AND name IN ('vFCPUFDest', 'vICMSUFDest', 'vICMSUFRemet')
ORDER BY name;
GO

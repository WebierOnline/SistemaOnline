-- ============================================================
-- Adiciona colunas de totais ICMS60 (ST Retido anteriormente)
-- na tabela cabecalho EntradaEstoque.
--
-- Esses valores sao a soma dos campos por item (ICMS60) de
-- EntradaEstoqueItens, acumulados no momento da importacao XML.
--
--   vBCSTRetTotal      : SUM(vBCSTRet)       - Base de calculo ST retido
--   vICMSSTRetTotal    : SUM(vICMSSTRet)     - Valor do ICMS ST retido
--   vICMSSubstitutoTotal: SUM(vICMSSubstituto)- Valor ICMS do substituto
-- ============================================================

IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('EntradaEstoque') AND name = 'vBCSTRetTotal'
)
    ALTER TABLE EntradaEstoque ADD vBCSTRetTotal DECIMAL(15,2) NOT NULL DEFAULT 0;
GO

IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('EntradaEstoque') AND name = 'vICMSSTRetTotal'
)
    ALTER TABLE EntradaEstoque ADD vICMSSTRetTotal DECIMAL(15,2) NOT NULL DEFAULT 0;
GO

IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('EntradaEstoque') AND name = 'vICMSSubstitutoTotal'
)
    ALTER TABLE EntradaEstoque ADD vICMSSubstitutoTotal DECIMAL(15,2) NOT NULL DEFAULT 0;
GO

-- Verificar colunas criadas
SELECT name, max_length, is_nullable
FROM sys.columns
WHERE object_id = OBJECT_ID('EntradaEstoque')
  AND name IN ('vBCSTRetTotal', 'vICMSSTRetTotal', 'vICMSSubstitutoTotal')
ORDER BY name;
GO

-- =============================================================================
-- Fix_IS_uTrib_qTrib_NF396.sql
-- Corrige dois problemas relacionados ao IS que causaram rejeicao da NF 396:
--   1. tbISClassTrib: uTrib_IS NULL para registros ad valorem (tipo=1)
--   2. NotaFiscalItens NF 396: IS_qUnid=0 e uTrib_IS vazio
--
-- Executar ANTES de reabrir/retransmitir a NF 396.
-- =============================================================================

-- PASSO 1: Garantir uTrib_IS em tbISClassTrib para tipo ad valorem (1)
-- Para calculo ad valorem, qTrib = quantidade comercial e uTrib = unidade comercial.
-- 'UN' (unidade) e valido e o mais comum; ajuste manualmente para casos especificos.
UPDATE tbISClassTrib
SET uTrib_IS = 'UN'
WHERE (uTrib_IS IS NULL OR LTRIM(RTRIM(CAST(uTrib_IS AS VARCHAR(20)))) = '')
  AND tipo_calculo_is = 1;

-- Verificar resultado:
-- SELECT cClassTrib_IS, tipo_calculo_is, uTrib_IS FROM tbISClassTrib ORDER BY cClassTrib_IS;


-- PASSO 2: Corrigir NotaFiscalItens da NF 396
-- IS_qUnid estava 0 (foi zerado incorretamente no Calcular_IS).
-- uTrib_IS estava vazio (tbISClassTrib nao tinha uTrib_IS quando a NF foi criada).
UPDATE nfi
SET
    nfi.IS_qUnid  = nfi.QuantidadeComercial,
    nfi.uTrib_IS  = COALESCE(
                        NULLIF(LTRIM(RTRIM(ISNULL(t.uTrib_IS, ''))), ''),
                        'UN'
                    )
FROM NotaFiscalItens nfi
OUTER APPLY (
    SELECT TOP 1 uTrib_IS
    FROM tbISClassTrib t2
    WHERE t2.cClassTrib_IS = nfi.cClassTrib_IS
    ORDER BY
        CASE WHEN GETDATE() BETWEEN t2.dIniVig AND t2.dFimVig THEN 0 ELSE 1 END ASC,
        t2.dFimVig DESC
) t
WHERE nfi.CodigoNota = (
    SELECT TOP 1 CodigoNota
    FROM NotaFiscal
    WHERE NumeroNota = 396
      AND SerieNF LIKE '%2%'
)
  AND nfi.IS_qUnid = 0
  AND nfi.IS_tipo_calculo > 0
  AND nfi.QuantidadeComercial > 0;

-- Verificar resultado:
-- SELECT ITEM, CodigoProduto, QuantidadeComercial, IS_tipo_calculo, IS_qUnid, uTrib_IS
-- FROM NotaFiscalItens
-- WHERE CodigoNota = (SELECT TOP 1 CodigoNota FROM NotaFiscal WHERE NumeroNota = 396 AND SerieNF LIKE '%2%')
-- ORDER BY ITEM;

-- Corrige aliquotas da reforma (IBS/CBS) zeradas nos itens de uma nota ainda em digitacao,
-- normalmente notas criadas antes da atualizacao. Ajuste @CodigoNota.
-- Usa os valores de referencia do ano-teste 2026: IBS UF 0,10 / IBS Mun 0,00 / CBS 0,90.
-- Recalcula IBS_vIBSUF / IBS_vIBSMun / IBS_vIBS / CBS_vCBS a partir das BCs ja gravadas.
-- So mexe em item com IBS_vBC > 0 (item que realmente entra no grupo IBSCBS). Idempotente.
SET NOCOUNT ON;
DECLARE @CodigoNota INT = 275;   -- <<< AJUSTE

UPDATE ni
SET ni.IBS_UFpAliq  = CASE WHEN ni.IBS_UFpAliq  IS NULL OR ni.IBS_UFpAliq  = 0 THEN 0.10 ELSE ni.IBS_UFpAliq  END,
    ni.IBS_MunpAliq = ISNULL(ni.IBS_MunpAliq, 0),
    ni.CBS_pAliq    = CASE WHEN ni.CBS_pAliq    IS NULL OR ni.CBS_pAliq    = 0 THEN 0.90 ELSE ni.CBS_pAliq    END
FROM NotaFiscalItens ni
WHERE ni.CodigoNota = @CodigoNota AND ni.IBS_vBC > 0;

UPDATE ni
SET ni.IBS_vIBSUF  = CAST(ROUND(ni.IBS_vBC * ni.IBS_UFpAliq  * (1 - ISNULL(ni.IBS_pRed,0)/100.0) / 100.0, 2) AS DECIMAL(15,2)),
    ni.IBS_vIBSMun = CAST(ROUND(ni.IBS_vBC * ni.IBS_MunpAliq * (1 - ISNULL(ni.IBS_pRed,0)/100.0) / 100.0, 2) AS DECIMAL(15,2)),
    ni.CBS_vCBS    = CAST(ROUND(ni.CBS_vBC * ni.CBS_pAliq    * (1 - ISNULL(ni.CBS_pRed,0)/100.0) / 100.0, 2) AS DECIMAL(15,2))
FROM NotaFiscalItens ni
WHERE ni.CodigoNota = @CodigoNota AND ni.IBS_vBC > 0;

UPDATE NotaFiscalItens
SET IBS_vIBS = ISNULL(IBS_vIBSUF,0) + ISNULL(IBS_vIBSMun,0)
WHERE CodigoNota = @CodigoNota;

PRINT 'Itens corrigidos:';
SELECT ITEM, IBS_vBC, IBS_UFpAliq, IBS_MunpAliq, IBS_vIBSUF, IBS_vIBSMun, IBS_vIBS, CBS_vBC, CBS_pAliq, CBS_vCBS
FROM NotaFiscalItens WHERE CodigoNota = @CodigoNota ORDER BY ITEM;

-- Preenche com o padrao (tributacao integral, sem IS) os produtos cujo NCM nao foi encontrado na
-- tbNCM mesmo apos a importacao completa da Tabela IBPT + tbNCM_preencher_IBSCBS.sql.
-- Mesmo escopo da consulta de diagnostico usada pra achar esses casos (motivo = "NCM nao encontrado
-- na tbNCM") - so mexe em quem realmente nao tem correspondencia, nao sobrescreve nada que ja bateu.

UPDATE p
SET p.cClassTrib = '000001',
    p.IBSCBSCST = '000',
    p.cClassTrib_IS = NULL,
    p.tipo_calculo_is = 0,
    p.ISCST = '00',
    p.fator_conversao_IS = 1.000
FROM produtos p
LEFT JOIN tbNCM n ON RTRIM(LTRIM(p.ncm)) = n.NCM
WHERE (p.cClassTrib IS NULL OR p.cClassTrib = '' OR p.IBSCBSCST IS NULL OR p.IBSCBSCST <> LEFT(p.cClassTrib, 3))
  AND p.ncm IS NOT NULL AND RTRIM(LTRIM(p.ncm)) <> ''
  AND n.NCM IS NULL;

-- Conferencia
SELECT p.codigo, p.descricao, p.ncm, p.cClassTrib, p.IBSCBSCST
FROM produtos p
LEFT JOIN tbNCM n ON RTRIM(LTRIM(p.ncm)) = n.NCM
WHERE (p.cClassTrib IS NULL OR p.cClassTrib = '')
   OR (p.IBSCBSCST IS NULL OR p.IBSCBSCST <> LEFT(p.cClassTrib, 3))
ORDER BY p.codigo;

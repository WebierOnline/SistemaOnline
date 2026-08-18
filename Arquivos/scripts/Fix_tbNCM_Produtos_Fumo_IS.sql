-- Fix pontual: NCM do capitulo 24 (fumo/cigarros e derivados) nunca tinha regra de
-- Imposto Seletivo em tbNCM_preencher_IBSCBS.sql (script rodado antes desse gap ser
-- corrigido) - produto com esse NCM ficava com cClassTrib_IS/tipo_calculo_is vazios
-- e saia sem o grupo <IS> na NFCe. Rodar em clientes que ja rodaram o manifesto ate
-- a posicao 061 (tbNCM_preencher_IBSCBS.sql) ANTES dessa correcao existir.
--
-- Idempotente - pode rodar de novo sem problema.

-- 1) DIAGNOSTICO - produtos de fumo (cap. 24) sem IS aplicado hoje
SELECT p.codigo, p.DESCRICAO, p.NCM, p.cClassTrib_IS, p.tipo_calculo_is
FROM produtos p
WHERE p.NCM LIKE '24%';

-- 2) Corrige a tbNCM (mesma regra adicionada em tbNCM_preencher_IBSCBS.sql)
UPDATE tbNCM SET
    cClassTrib_IS = '900040',
    tipo_calculo_is = 3
WHERE NCM LIKE '24%';

-- 3) Resincroniza os produtos com NCM de fumo a partir da tbNCM (mesma logica de
--    Popular_Produtos_ClassTrib_ViaNCM.sql, mas escopada so pro capitulo 24 pra nao
--    reprocessar produtos que ja estao corretos)
UPDATE p
SET p.cClassTrib = n.cClassTrib_IBS,
    p.IBSCBSCST = LEFT(n.cClassTrib_IBS, 3),
    p.cClassTrib_IS = NULLIF(n.cClassTrib_IS, ''),
    p.tipo_calculo_is = n.tipo_calculo_is,
    p.ISCST = CASE WHEN n.cClassTrib_IS <> '' THEN '01' ELSE '00' END,
    p.fator_conversao_IS = ISNULL(p.fator_conversao_IS, 1)
FROM produtos p
INNER JOIN tbNCM n ON RTRIM(LTRIM(p.ncm)) = n.NCM
WHERE p.NCM LIKE '24%'
  AND n.cClassTrib_IBS <> '';

-- 4) CONFERENCIA - deve vir tudo preenchido (cClassTrib_IS = '900040', tipo_calculo_is = 3)
SELECT p.codigo, p.DESCRICAO, p.NCM, p.cClassTrib_IS, p.tipo_calculo_is
FROM produtos p
WHERE p.NCM LIKE '24%';

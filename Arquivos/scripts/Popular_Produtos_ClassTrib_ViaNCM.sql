-- Resincroniza produtos.cClassTrib / IBSCBSCST / cClassTrib_IS / tipo_calculo_is / ISCST
-- a partir do NCM do produto, usando a classificacao ja existente em tbNCM (populada pelo script
-- tbNCM_preencher_IBSCBS.sql).
--
-- IMPORTANTE: diferente da 1a versao deste script, este SEMPRE resincroniza a partir da tbNCM,
-- mesmo em produto que ja tem cClassTrib preenchido - varios clientes tiveram cClassTrib='000001'
-- preenchido em massa manualmente em algum momento (tributacao integral pra tudo), o que deixa
-- errados os produtos que deveriam ser cesta basica/GLP/etc (cClassTrib diferente de '000001') e
-- deixa os campos de IS (cClassTrib_IS/tipo_calculo_is/ISCST) sempre vazios, ja que esse
-- preenchimento manual nunca tocou a parte de Imposto Seletivo.
-- Continua idempotente (pode rodar de novo sempre que a tbNCM for atualizada) e nao mexe em
-- produto sem correspondencia de NCM na tbNCM.
--
-- Regras assumidas (ja confirmadas em teste real, NFCe autorizada pela SEFAZ em producao):
--   1) IBSCBSCST = primeiros 3 digitos de cClassTrib (padrao observado em toda a tabela oficial
--      TbIBSCBSClassTrib: cClassTrib '000001' -> CST '000', '400001' -> CST '400', etc.)
--   2) ISCST = '01' (Saida Tributada) quando o produto tem cClassTrib_IS aplicavel,
--      '00' (Nao incidencia) quando nao tem - mesmo padrao usado em tbISClassTrib.ISCST.
--   3) fator_conversao_IS: NAO vem da tbNCM (e especifico de embalagem, nao de NCM) - mantem o que
--      ja estiver no produto; se nulo, assume 1 (o calculo em VB6 ja tem esse fallback). Produtos
--      sujeitos a IS especifico/misto (bebidas, por ex.) podem precisar de ajuste fino manual desse
--      fator conforme a embalagem real.

-- 1) DIAGNOSTICO - rodar antes de aplicar, pra saber o tamanho do problema
SELECT COUNT(*) AS produtos_sem_classtrib_hoje
FROM produtos
WHERE (cClassTrib IS NULL OR cClassTrib = '');

SELECT COUNT(*) AS ja_preenchido_mas_ibscbscst_desatualizado  -- cClassTrib existe mas nao bate com LEFT(cClassTrib,3)
FROM produtos
WHERE cClassTrib IS NOT NULL AND cClassTrib <> ''
  AND (IBSCBSCST IS NULL OR IBSCBSCST <> LEFT(cClassTrib, 3));

SELECT COUNT(*) AS serao_atualizados_por_este_script  -- inclui os 2 casos acima, sempre que tiver NCM na tbNCM
FROM produtos p
INNER JOIN tbNCM n ON RTRIM(LTRIM(p.ncm)) = n.NCM
WHERE n.cClassTrib_IBS <> ''
  AND (
        p.cClassTrib IS NULL OR p.cClassTrib <> n.cClassTrib_IBS
     OR p.IBSCBSCST IS NULL OR p.IBSCBSCST <> LEFT(n.cClassTrib_IBS, 3)
     OR p.cClassTrib_IS IS NULL OR p.cClassTrib_IS <> NULLIF(n.cClassTrib_IS, '')
     OR ISNULL(p.tipo_calculo_is, -1) <> n.tipo_calculo_is
      );

SELECT COUNT(*) AS sem_correspondencia_na_tbNCM  -- precisam de tratamento manual (NCM invalido/nao cadastrado em tbNCM)
FROM produtos p
LEFT JOIN tbNCM n ON RTRIM(LTRIM(p.ncm)) = n.NCM
WHERE (n.NCM IS NULL OR n.cClassTrib_IBS = '')
  AND (p.cClassTrib IS NULL OR p.cClassTrib = '');

-- 2) ATUALIZACAO EM MASSA - sempre resincroniza com a tbNCM (sobrescreve cClassTrib/IBSCBSCST/
--    cClassTrib_IS/tipo_calculo_is/ISCST existentes, inclusive os preenchidos manualmente em massa)
UPDATE p
SET p.cClassTrib = n.cClassTrib_IBS,
    p.IBSCBSCST = LEFT(n.cClassTrib_IBS, 3),
    p.cClassTrib_IS = NULLIF(n.cClassTrib_IS, ''),
    p.tipo_calculo_is = n.tipo_calculo_is,
    p.ISCST = CASE WHEN n.cClassTrib_IS <> '' THEN '01' ELSE '00' END,
    p.fator_conversao_IS = ISNULL(p.fator_conversao_IS, 1)
FROM produtos p
INNER JOIN tbNCM n ON RTRIM(LTRIM(p.ncm)) = n.NCM
WHERE n.cClassTrib_IBS <> '';

-- 3) CONFERENCIA - deve bater com "sem_correspondencia_na_tbNCM" do diagnostico acima
SELECT COUNT(*) AS ainda_sem_classtrib_apos_rodar
FROM produtos
WHERE (cClassTrib IS NULL OR cClassTrib = '');

SELECT COUNT(*) AS ainda_desatualizado_apos_rodar  -- deve vir 0 (ou so os "sem correspondencia")
FROM produtos
WHERE cClassTrib IS NOT NULL AND cClassTrib <> ''
  AND (IBSCBSCST IS NULL OR IBSCBSCST <> LEFT(cClassTrib, 3));

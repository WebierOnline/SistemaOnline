-- "Linha 1" (fase de testes, enquanto o governo nao define os CST/aliquotas
-- definitivos da reforma tributaria IBS/CBS/IS, 2026-10-27): o script
-- Popular_Produtos_ClassTrib_ViaNCM.sql (067 no manifesto) busca o cClassTrib
-- especifico de cada produto pelo NCM na tbNCM - ja deu erro de transmissao
-- (SEFAZ rejeitando) em alguns clientes, porque certos NCMs caem num
-- IBSCBSCST diferente de '000' que ainda nao esta sendo aceito de forma
-- confiavel. Enquanto isso nao se resolve, usar esse script aqui: fixa
-- TODOS os produtos em tributacao integral generica (cClassTrib='000001',
-- IBSCBSCST='000'), sem excecao por NCM - mesma logica de fallback que o
-- script 068 (Fix_Produtos_ClassTrib_SemNCM_NaTbNCM.sql) ja usa pros
-- produtos sem NCM na tbNCM, so que aplicada pra tabela inteira.
--
-- O 067 foi retirado do manifesto de proposito nesta fase - NAO rodar os
-- dois juntos (o 067 sobrescreveria o valor fixo com o valor por NCM de
-- novo). Quando o governo definir os CST/aliquotas definitivos ("linha 2"),
-- volta a usar o 067 (esse script aqui deixa de ser necessario a partir dai).
--
-- Idempotente - pode rodar de novo sem problema.

-- 1) DIAGNOSTICO - quantos produtos vao ser alterados
SELECT COUNT(*) AS produtos_fora_do_padrao_fixo
FROM produtos
WHERE ISNULL(cClassTrib, '') <> '000001' OR ISNULL(IBSCBSCST, '') <> '000';

-- 2) ATUALIZACAO EM MASSA - fixa TODOS os produtos (sobrescreve qualquer
--    valor anterior, inclusive o que o script 067 ja tenha preenchido por NCM)
UPDATE produtos
SET cClassTrib = '000001',
    IBSCBSCST = '000',
    cClassTrib_IS = NULL,
    tipo_calculo_is = 0,
    ISCST = '00',
    fator_conversao_IS = 1.000;

-- 3) CONFERENCIA - deve vir 0
SELECT COUNT(*) AS produtos_fora_do_padrao_apos_rodar
FROM produtos
WHERE ISNULL(cClassTrib, '') <> '000001' OR ISNULL(IBSCBSCST, '') <> '000';

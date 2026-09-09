-- 1. LIMPEZA E REGRA GERAL (Aliquota Padrao)
-- Define todos como IBS Padrao (000001) e sem Imposto Seletivo
UPDATE tbNCM SET 
    cClassTrib_IBS = '000001', 
    cClassTrib_IS = '', 
    tipo_calculo_is = 0;

-- 2. CESTA BASICA NACIONAL (Art. 125, Anexo I da LC 214/2025 - reducao de 100% = aliquota zero)
-- cClassTrib 200003 = "Vendas de produtos destinados a alimentacao humana (Anexo I)", CST 200.
-- CORRIGIDO EM 2026-09-01: estava usando 400001 (CST 400 = isencao de transporte publico de
-- passageiros - nada a ver com alimento), o que a SEFAZ rejeitou em teste real (NCM 10063021 - arroz).
-- NCMs de Carnes (0201-0204), Peixes (0302-0305), Leite (0401), Feijao (0713), Arroz (1006)
UPDATE tbNCM SET 
    cClassTrib_IBS = '200003' 
WHERE NCM LIKE '0201%' OR NCM LIKE '0202%' OR NCM LIKE '0203%' OR NCM LIKE '0204%'
   OR NCM LIKE '0302%' OR NCM LIKE '0303%' OR NCM LIKE '0304%' OR NCM LIKE '0305%'
   OR NCM LIKE '0401%' OR NCM LIKE '0713%' OR NCM LIKE '1006%';

-- 3. HIGIENE PESSOAL E LIMPEZA (Art. 136, Anexo VIII da LC 214/2025)
-- cClassTrib 200035 = "Fornecimento dos produtos de higiene pessoal e limpeza relacionados no Anexo VIII".
-- CORRIGIDO EM 2026-09-01: estava usando 200001 (que na verdade e transporte de bens ate zona de
-- processamento de exportacao - nao tem nada a ver com higiene/limpeza).
-- Higiene (3401), Limpeza (3402)
UPDATE tbNCM SET 
    cClassTrib_IBS = '200035'
WHERE NCM LIKE '3401%' OR NCM LIKE '3402%';

-- 3b. HORTIFRUTI FRESCO/CONGELADO NAO COZIDO (Art. 148, Anexo XV da LC 214/2025)
-- cClassTrib 200014 = horticolas, frutas e ovos, desde que nao cozidos. NCM 0710 e congelado mas
-- nao cozido na maioria dos casos (excecao: trufas, NCM 0710.80.00, fora do Anexo XV - raro nesse
-- ramo de negocio). CONFIANCA MENOR que os itens 2 e 3 acima - revisar se aparecer produto de NCM
-- 0710 dando problema na SEFAZ.
UPDATE tbNCM SET 
    cClassTrib_IBS = '200014'
WHERE NCM LIKE '0710%';

-- 4. IMPOSTO SELETIVO - BEBIDAS ALCOOLICAS (Ad Valorem %)
-- Cervejas (2203), Vinhos (2204), Aguardentes (2208)
UPDATE tbNCM SET 
    cClassTrib_IS = '900001', 
    tipo_calculo_is = 1 
WHERE NCM LIKE '2203%' OR NCM LIKE '2204%' OR NCM LIKE '2208%';

-- 5. IMPOSTO SELETIVO - BEBIDAS ACUCARADAS (Ad Rem R$/Litro)
-- Refrigerantes, Refrescos, Isotonicos (2202)
UPDATE tbNCM SET 
    cClassTrib_IS = '900010', 
    tipo_calculo_is = 2 
WHERE NCM LIKE '2202%';

-- 6. IMPOSTO SELETIVO - PRODUTOS DO FUMO (Misto: % + R$/unidade)
-- Cigarros e derivados do tabaco (capitulo 24 da NCM)
UPDATE tbNCM SET
    cClassTrib_IS = '900040',
    tipo_calculo_is = 3
WHERE NCM LIKE '24%';

-- 7. COMBUSTIVEIS - REGIME MONOFASICO (Gas GLP)
-- NCM do GLP (27111910) - cClassTrib 620006 NAO CONFIRMADO em fonte oficial (busca nao achou
-- documentacao especifica). Mantido como estava; validar antes de confiar em cliente que vende GLP.
UPDATE tbNCM SET
    cClassTrib_IBS = '620006',
    cClassTrib_IS = '',
    tipo_calculo_is = 2
WHERE NCM = '27111910';

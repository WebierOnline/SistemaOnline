-- 1. LIMPEZA E REGRA GERAL (Alíquota Padrão)
-- Define todos como IBS Padrão (000001) e sem Imposto Seletivo
UPDATE tbNCM SET 
    cClassTrib_IBS = '000001', 
    cClassTrib_IS = '', 
    tipo_calculo_is = 0;

-- 2. CESTA BÁSICA NACIONAL (Alíquota Zero)
-- NCMs de Carnes (0201-0204), Peixes (0302-0305), Leite (0401), Feijão (0713), Arroz (1006)
UPDATE tbNCM SET 
    cClassTrib_IBS = '400001' 
WHERE NCM LIKE '0201%' OR NCM LIKE '0202%' OR NCM LIKE '0203%' OR NCM LIKE '0204%'
   OR NCM LIKE '0302%' OR NCM LIKE '0303%' OR NCM LIKE '0304%' OR NCM LIKE '0305%'
   OR NCM LIKE '0401%' OR NCM LIKE '0713%' OR NCM LIKE '1006%';

-- 3. CESTA BÁSICA COM REDUÇÃO (Alíquota 60%)
-- Higiene (3401), Limpeza (3402), Hortifrúti processado (0710)
UPDATE tbNCM SET 
    cClassTrib_IBS = '200001'
WHERE NCM LIKE '3401%' OR NCM LIKE '3402%' OR NCM LIKE '0710%';

-- 4. IMPOSTO SELETIVO - BEBIDAS ALCOÓLICAS (Ad Valorem %)
-- Cervejas (2203), Vinhos (2204), Aguardentes (2208)
UPDATE tbNCM SET 
    cClassTrib_IS = '900001', 
    tipo_calculo_is = 1 
WHERE NCM LIKE '2203%' OR NCM LIKE '2204%' OR NCM LIKE '2208%';

-- 5. IMPOSTO SELETIVO - BEBIDAS AÇUCARADAS (Ad Rem R$/Litro)
-- Refrigerantes, Refrescos, Isotônicos (2202)
UPDATE tbNCM SET 
    cClassTrib_IS = '900010', 
    tipo_calculo_is = 2 
WHERE NCM LIKE '2202%';

-- 6. COMBUSTÍVEIS - REGIME MONOFÁSICO (Gás GLP)
-- NCM do GLP (27111910) - Usa cClassTrib 620006 e cálculo Ad Rem
UPDATE tbNCM SET 
    cClassTrib_IBS = '620006', 
    cClassTrib_IS = '',
    tipo_calculo_is = 2
WHERE NCM = '27111910';
-- Marca CPF = 1 para todos os registros de empresas_desbloueio cujo
-- campo CNPJ NAO esta no formato de CNPJ (##.###.###/####-##).
-- CNPJ sempre tem uma barra "/" (ex: 63.139.733/0001-21); CPF nao tem
-- (ex: 770.235.443-72) -- entao "nao ter barra" identifica CPF com seguranca.

-- 1) Preview: confira quem vai ser afetado antes de rodar o UPDATE
SELECT Codigo, Fantasia, Razao, CNPJ, CPF AS CPF_Atual
FROM empresas_desbloueio
WHERE CNPJ NOT LIKE '%/%'
  AND CNPJ IS NOT NULL
  AND CNPJ <> '';

-- 2) UPDATE propriamente dito
UPDATE empresas_desbloueio
SET CPF = 1
WHERE CNPJ NOT LIKE '%/%'
  AND CNPJ IS NOT NULL
  AND CNPJ <> '';

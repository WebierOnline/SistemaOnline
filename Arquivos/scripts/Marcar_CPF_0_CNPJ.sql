-- Marca CPF = 0 para todos os registros de empresas_desbloueio cujo
-- campo CNPJ ESTA no formato de CNPJ (##.###.###/####-##), ou seja,
-- contem uma barra "/" (ex: 63.139.733/0001-21).

-- 1) Preview: confira quem vai ser afetado antes de rodar o UPDATE
SELECT Codigo, Fantasia, Razao, CNPJ, CPF AS CPF_Atual
FROM empresas_desbloueio
WHERE CNPJ LIKE '%/%';

-- 2) UPDATE propriamente dito
UPDATE empresas_desbloueio
SET CPF = 0
WHERE CNPJ LIKE '%/%';

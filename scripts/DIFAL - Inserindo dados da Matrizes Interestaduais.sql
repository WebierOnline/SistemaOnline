-- Sul e Sudeste (exceto ES) enviando para Norte, Nordeste, Centro-Oeste e ES (7%)
INSERT INTO TribMatrizInterestadual (UF_Origem, UF_Destino, AliquotaInterestadual)
SELECT Origem, Destino, 7.00
FROM (VALUES ('SP'),('MG'),('RJ'),('PR'),('RS'),('SC')) AS O(Origem)
CROSS JOIN (VALUES ('AC'),('AL'),('AM'),('AP'),('BA'),('CE'),('DF'),('ES'),('GO'),('MA'),('MT'),('MS'),('PA'),('PB'),('PE'),('PI'),('RN'),('RO'),('RR'),('SE'),('TO')) AS D(Destino);

-- Demais operações e Interestaduais entre estados da mesma região (12%)
-- O comando abaixo preenche o que falta com 12%
INSERT INTO TribMatrizInterestadual (UF_Origem, UF_Destino, AliquotaInterestadual)
SELECT O.UF, D.UF, 12.00
FROM (VALUES ('AC'),('AL'),('AM'),('AP'),('BA'),('CE'),('DF'),('ES'),('GO'),('MA'),('MT'),('MS'),('MG'),('PA'),('PB'),('PR'),('PE'),('PI'),('RJ'),('RN'),('RS'),('RO'),('RR'),('SC'),('SP'),('SE'),('TO')) AS O(UF)
CROSS JOIN (VALUES ('AC'),('AL'),('AM'),('AP'),('BA'),('CE'),('DF'),('ES'),('GO'),('MA'),('MT'),('MS'),('MG'),('PA'),('PB'),('PR'),('PE'),('PI'),('RJ'),('RN'),('RS'),('RO'),('RR'),('SC'),('SP'),('SE'),('TO')) AS D(UF)
WHERE O.UF <> D.UF
AND NOT EXISTS (SELECT 1 FROM TribMatrizInterestadual M WHERE M.UF_Origem = O.UF AND M.UF_Destino = D.UF);

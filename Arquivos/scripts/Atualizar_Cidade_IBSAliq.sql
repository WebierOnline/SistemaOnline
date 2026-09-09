-- Preenche as aliquotas de referencia do IBS (ano-teste 2026: IBS UF 0,10% / IBS Mun 0,00%).
-- Pega tanto linhas NULL quanto = 0 - a coluna foi criada com DEFAULT 0, entao dependendo da
-- versao do SQL Server / de reimportacao de cidades as linhas podem estar 0 e o "IS NULL"
-- sozinho deixava passar (SEFAZ rejeita: "Aliquota do IBS da UF invalida"). Idempotente.
UPDATE Cidade SET IBSUFpAliq  = 0.10 WHERE IBSUFpAliq  IS NULL OR IBSUFpAliq  = 0;
GO
UPDATE Cidade SET IBSMunpAliq = 0.00 WHERE IBSMunpAliq IS NULL;
GO
PRINT 'IBSUFpAliq / IBSMunpAliq de referencia (2026) aplicadas.';
GO

-- Juros e multa por atraso (Configuracao Geral > Juros):
--   JUROS_MES = 1,00 (% ao mes) e JUROS_DIA = 0,0333 (% ao dia) para todos os clientes;
--   MULTA (nova linha, % cobrada uma vez sobre a parcela vencida) = 0,00 se ainda nao existir.
-- Valores com virgula, como a tela grava. Idempotente.
IF EXISTS (SELECT 1 FROM configuracao WHERE config_nome = 'JUROS_MES')
    UPDATE configuracao SET config_valor = '1,00' WHERE config_nome = 'JUROS_MES'
ELSE
    INSERT INTO configuracao (config_nome, config_valor) VALUES ('JUROS_MES', '1,00')
GO

IF EXISTS (SELECT 1 FROM configuracao WHERE config_nome = 'JUROS_DIA')
    UPDATE configuracao SET config_valor = '0,0333' WHERE config_nome = 'JUROS_DIA'
ELSE
    INSERT INTO configuracao (config_nome, config_valor) VALUES ('JUROS_DIA', '0,0333')
GO

IF NOT EXISTS (SELECT 1 FROM configuracao WHERE config_nome = 'MULTA')
    INSERT INTO configuracao (config_nome, config_valor) VALUES ('MULTA', '0,00')
GO

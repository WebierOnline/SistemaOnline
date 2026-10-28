-- Configuracao_Geral.frm: novo checkbox chkBloqueioAPrazo (1a aba, GUIA: GERAL) - controla
-- se vendas a prazo ficam bloqueadas.
-- A tabela configuracao e generica (config_nome/config_valor), nao precisa de ALTER TABLE -
-- so falta a LINHA default ('0' = NAO/false, desbloqueado).
IF NOT EXISTS (SELECT 1 FROM configuracao WHERE config_nome = 'BLOQUEIO_APRAZO')
    INSERT INTO configuracao (config_nome, config_valor) VALUES ('BLOQUEIO_APRAZO', '0');
GO

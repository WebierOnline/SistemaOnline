-- Configuracao_Geral.frm: novo checkbox chkExibirTributos (guia PDV) - EXIBIR_TRIBUTOS ('1' = exibe, '0' = nao exibe).
-- A tabela configuracao e generica (config_nome/config_valor); so falta a LINHA, default '0' (false) para todos.
IF NOT EXISTS (SELECT 1 FROM configuracao WHERE config_nome = 'EXIBIR_TRIBUTOS')
    INSERT INTO configuracao (config_nome, config_valor) VALUES ('EXIBIR_TRIBUTOS', '0');
GO

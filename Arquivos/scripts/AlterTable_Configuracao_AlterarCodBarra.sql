-- Configuracao_Geral.frm: novo combo cboAlterarCodigo (1a aba, GUIA: GERAL) - controla se
-- o usuario pode alterar o codigo de barra de um produto ja cadastrado.
-- A tabela configuracao e generica (config_nome/config_valor), nao precisa de ALTER TABLE -
-- so falta a LINHA default ('1' = SIM/true).
IF NOT EXISTS (SELECT 1 FROM configuracao WHERE config_nome = 'ALTERAR_CODBARRA')
    INSERT INTO configuracao (config_nome, config_valor) VALUES ('ALTERAR_CODBARRA', '1');
GO

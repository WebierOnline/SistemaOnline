-- Configuracao_Geral.frm: novo combo cboHabilitarComp (Habilitar Compartibilidade) - controla
-- se o PDV deve exibir compatibilidades de veiculos (TIPO_OS = Automoveis/Motocicletas e
-- TIPO_EMPRESA = 5 "Autopeca/Motopeca") no lstBusca.
-- A tabela configuracao e generica (config_nome/config_valor), nao precisa de ALTER TABLE -
-- so falta a LINHA default ('0' = NAO/false).
IF NOT EXISTS (SELECT 1 FROM configuracao WHERE config_nome = 'HABILITAR_COMP')
    INSERT INTO configuracao (config_nome, config_valor) VALUES ('HABILITAR_COMP', '0');
GO

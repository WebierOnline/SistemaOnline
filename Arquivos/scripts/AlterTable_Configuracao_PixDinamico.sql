-- PDV: forma de pagamento rapida (imagens/F8-F12). PIX_DINAMICO: '1' = a imagem Pix escolhe "11 - PIX DINAMICO"; '0' = "10 - PIX ESTATICO".
-- A tabela configuracao e generica (config_nome/config_valor); so falta a LINHA, default '0'.
IF NOT EXISTS (SELECT 1 FROM configuracao WHERE config_nome = 'PIX_DINAMICO')
    INSERT INTO configuracao (config_nome, config_valor) VALUES ('PIX_DINAMICO', '0');
GO

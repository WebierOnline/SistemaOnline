-- Produtos_Cadastro.frm: GerarProximoCodBarraAuto nao depende mais de MAX(COD_BARRA) da
-- tabela produtos - uma digitacao manual de 5 digitos (ex: codigo de balanca de
-- acougue/verdurao) "sequestrava" o proximo numero gerado, pulando pra uma faixa alta sem
-- necessidade real e se aproximando do limite de 99999. Agora o ultimo codigo gerado pelo
-- PROPRIO sistema fica isolado nessa linha de configuracao.
-- A tabela configuracao e generica (config_nome/config_valor), nao precisa de ALTER TABLE -
-- so falta a LINHA default, semeada com o maior codigo de 5 digitos ja usado hoje (pra nao
-- colidir com codigos ja existentes nos clientes que ja tem produtos cadastrados).
IF NOT EXISTS (SELECT 1 FROM configuracao WHERE config_nome = 'ULTIMO_CODBARRA_AUTO')
BEGIN
    DECLARE @UltimoCodBarra INT;
    SELECT @UltimoCodBarra = ISNULL(MAX(CAST(COD_BARRA AS INT)), 0)
    FROM produtos
    WHERE LEN(COD_BARRA) = 5 AND ISNUMERIC(COD_BARRA) = 1;

    INSERT INTO configuracao (config_nome, config_valor) VALUES ('ULTIMO_CODBARRA_AUTO', CAST(@UltimoCodBarra AS VARCHAR(10)));
END
GO

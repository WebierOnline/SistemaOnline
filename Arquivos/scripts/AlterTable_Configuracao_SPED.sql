-- SPED Fiscal (EFD ICMS/IPI) no envio mensal ao contador (OnlineCommerce /enviarsped + SpedGerador.exe).
-- SPED_ENVIO: 0 = nao envia; 1 = so o Bloco H do inventario de 31/12, junto com os XML de janeiro (o contador
--             pede em fevereiro); 2 = EFD ICMS/IPI completa todo mes (a de fevereiro com o Bloco H de 31/12);
--             3 (padrao) = os dois: a completa todo mes e, no envio de fevereiro, tambem o arquivo so com o Bloco H.
--             Cada arquivo vai num e-mail separado do dos XML.
-- SPED_COD_RECEITA: codigo de receita do ICMS no registro E116, da tabela da UF do cliente (o contador informa).
--             So e usado quando ha ICMS a recolher no mes (sem valor a recolher o E116 nao e gerado).
IF NOT EXISTS (SELECT 1 FROM configuracao WHERE config_nome = 'SPED_ENVIO')
    INSERT INTO configuracao (config_nome, config_valor) VALUES ('SPED_ENVIO', '3');
GO
IF NOT EXISTS (SELECT 1 FROM configuracao WHERE config_nome = 'SPED_COD_RECEITA')
    INSERT INTO configuracao (config_nome, config_valor) VALUES ('SPED_COD_RECEITA', '');
GO

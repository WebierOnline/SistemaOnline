-- 3 casas -> 2 casas decimais nas tabelas do modulo OS. Guardado contra tabela inexistente
-- (cliente que ativou OS mas cujo banco ainda nao tem as tabelas do modulo).
IF OBJECT_ID('OS_Servicos_Recapadora', 'U') IS NOT NULL
BEGIN
    ALTER TABLE OS_Servicos_Recapadora ALTER COLUMN Preco decimal(16,2);
    ALTER TABLE OS_Servicos_Recapadora ALTER COLUMN Desconto decimal(16,2);
    ALTER TABLE OS_Servicos_Recapadora ALTER COLUMN Subtotal decimal(16,2);
    ALTER TABLE OS_Servicos_Recapadora ALTER COLUMN Total decimal(16,2);
END
GO
IF OBJECT_ID('OS_servicos_Auto', 'U') IS NOT NULL
BEGIN
    ALTER TABLE OS_servicos_Auto ALTER COLUMN Preco decimal(16,2);
    ALTER TABLE OS_servicos_Auto ALTER COLUMN Desconto decimal(16,2);
    ALTER TABLE OS_servicos_Auto ALTER COLUMN Subtotal decimal(16,2);
    ALTER TABLE OS_servicos_Auto ALTER COLUMN Total decimal(16,2);
END
GO
IF OBJECT_ID('OS', 'U') IS NOT NULL
BEGIN
    ALTER TABLE OS ALTER COLUMN SUBTOTAL decimal(16,2);
    ALTER TABLE OS ALTER COLUMN VALOR_DESC decimal(16,2);
    ALTER TABLE OS ALTER COLUMN ValorDescReal decimal(16,2);
    ALTER TABLE OS ALTER COLUMN TOTAL decimal(16,2);
END
GO

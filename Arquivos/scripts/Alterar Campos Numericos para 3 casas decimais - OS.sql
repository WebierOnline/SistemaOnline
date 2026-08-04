ALTER TABLE OS_Servicos_Recapadora ALTER COLUMN Preco decimal(16,2);
ALTER TABLE OS_Servicos_Recapadora ALTER COLUMN Desconto decimal(16,2);
ALTER TABLE OS_Servicos_Recapadora ALTER COLUMN Subtotal decimal(16,2);
ALTER TABLE OS_Servicos_Recapadora ALTER COLUMN Total decimal(16,2);

ALTER TABLE OS_servicos_Auto ALTER COLUMN Preco decimal(16,2);
ALTER TABLE OS_servicos_Auto ALTER COLUMN Desconto decimal(16,2);
ALTER TABLE OS_servicos_Auto ALTER COLUMN Subtotal decimal(16,2);
ALTER TABLE OS_servicos_Auto ALTER COLUMN Total decimal(16,2);

ALTER TABLE OS ALTER COLUMN SUBTOTAL decimal(16,2);
ALTER TABLE OS ALTER COLUMN VALOR_DESC decimal(16,2);
ALTER TABLE OS ALTER COLUMN ValorDescReal decimal(16,2);
ALTER TABLE OS ALTER COLUMN TOTAL decimal(16,2);

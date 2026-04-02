-- Modalidade de determinacao da BC do ICMS por produto
-- 0 = Margem de Valor Agregado (%)
-- 1 = Pauta (Valor)
-- 2 = Preco Tabelado Maximo (Valor)
-- 3 = Valor da Operacao (padrao - mais comum)
IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('produtos') AND name = 'modBC'
)
BEGIN
    ALTER TABLE produtos ADD modBC TINYINT NOT NULL DEFAULT 3;
END

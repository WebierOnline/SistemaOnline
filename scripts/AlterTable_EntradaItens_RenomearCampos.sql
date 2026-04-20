-- Renomeia e ajusta tipos de 4 campos em produtos_entrada_itens
-- para ficar igual aos campos correspondentes em EntradaEstoqueItens

-- 1. Renomear CODIGO_PRODUTO -> CodigoProduto (int permanece igual)
EXEC sp_rename 'produtos_entrada_itens.CODIGO_PRODUTO', 'CodigoProduto', 'COLUMN';

-- 2. Renomear DESCRICAO -> NomeProduto e ajustar tipo
EXEC sp_rename 'produtos_entrada_itens.DESCRICAO', 'NomeProduto', 'COLUMN';
ALTER TABLE produtos_entrada_itens ALTER COLUMN NomeProduto text NULL;

-- 3. Renomear QUANT -> QuantidadeTributavel e ajustar tipo
EXEC sp_rename 'produtos_entrada_itens.QUANT', 'QuantidadeTributavel', 'COLUMN';
ALTER TABLE produtos_entrada_itens ALTER COLUMN QuantidadeTributavel decimal(15,3) NULL;

-- 4. EAN: sem renomear, apenas ajustar tipo nvarchar(100) -> varchar(14)
--    Precisa dropar o DEFAULT constraint antes de alterar o tipo
DECLARE @sql NVARCHAR(500)
SELECT @sql = 'ALTER TABLE produtos_entrada_itens DROP CONSTRAINT ' + dc.name
FROM sys.default_constraints dc
INNER JOIN sys.columns c ON dc.parent_object_id = c.object_id AND dc.parent_column_id = c.column_id
WHERE c.object_id = OBJECT_ID('produtos_entrada_itens') AND c.name = 'EAN'

IF @sql IS NOT NULL
    EXEC sp_executesql @sql

ALTER TABLE produtos_entrada_itens ALTER COLUMN EAN varchar(14) NULL;

-- Remove constraints DEFAULT de cada coluna, depois dropa as colunas
-- produtos_entrada_itens

DECLARE @sql NVARCHAR(500)
DECLARE @col NVARCHAR(100)

-- Loop para cada coluna a remover
DECLARE cols CURSOR FOR
    SELECT name FROM sys.columns
    WHERE object_id = OBJECT_ID('produtos_entrada_itens')
      AND name IN ('CUSTO','VALOR_VV','VALOR_VP','VALOR_AV','VALOR_AP',
                   'MARGEM_VV','MARGEM_VP','MARGEM_AV','MARGEM_AP')

OPEN cols
FETCH NEXT FROM cols INTO @col
WHILE @@FETCH_STATUS = 0
BEGIN
    SELECT @sql = 'ALTER TABLE produtos_entrada_itens DROP CONSTRAINT ' + dc.name
    FROM sys.default_constraints dc
    INNER JOIN sys.columns c ON dc.parent_object_id = c.object_id AND dc.parent_column_id = c.column_id
    WHERE c.object_id = OBJECT_ID('produtos_entrada_itens') AND c.name = @col

    IF @sql IS NOT NULL
    BEGIN
        EXEC sp_executesql @sql
        SET @sql = NULL
    END

    FETCH NEXT FROM cols INTO @col
END
CLOSE cols
DEALLOCATE cols

-- Agora dropa as colunas
ALTER TABLE produtos_entrada_itens
DROP COLUMN CUSTO, VALOR_VV, VALOR_VP, VALOR_AV, VALOR_AP,
            MARGEM_VV, MARGEM_VP, MARGEM_AV, MARGEM_AP;

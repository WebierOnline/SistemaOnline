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

-- Agora dropa as colunas (uma por vez, so se ainda existir)
IF EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('produtos_entrada_itens') AND name = 'CUSTO')
    ALTER TABLE produtos_entrada_itens DROP COLUMN CUSTO;
IF EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('produtos_entrada_itens') AND name = 'VALOR_VV')
    ALTER TABLE produtos_entrada_itens DROP COLUMN VALOR_VV;
IF EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('produtos_entrada_itens') AND name = 'VALOR_VP')
    ALTER TABLE produtos_entrada_itens DROP COLUMN VALOR_VP;
IF EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('produtos_entrada_itens') AND name = 'VALOR_AV')
    ALTER TABLE produtos_entrada_itens DROP COLUMN VALOR_AV;
IF EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('produtos_entrada_itens') AND name = 'VALOR_AP')
    ALTER TABLE produtos_entrada_itens DROP COLUMN VALOR_AP;
IF EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('produtos_entrada_itens') AND name = 'MARGEM_VV')
    ALTER TABLE produtos_entrada_itens DROP COLUMN MARGEM_VV;
IF EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('produtos_entrada_itens') AND name = 'MARGEM_VP')
    ALTER TABLE produtos_entrada_itens DROP COLUMN MARGEM_VP;
IF EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('produtos_entrada_itens') AND name = 'MARGEM_AV')
    ALTER TABLE produtos_entrada_itens DROP COLUMN MARGEM_AV;
IF EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('produtos_entrada_itens') AND name = 'MARGEM_AP')
    ALTER TABLE produtos_entrada_itens DROP COLUMN MARGEM_AP;

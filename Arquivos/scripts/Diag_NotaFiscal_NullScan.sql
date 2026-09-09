-- Varredura de colunas NULL numa nota especifica (cabecalho + itens).
-- Ajuste o @CodigoNota abaixo (e o PK: CodigoNota, nao o numero fiscal).
-- Compativel com SQL Server 2008.
SET NOCOUNT ON;

DECLARE @CodigoNota INT = 275;   -- <<< AJUSTE AQUI

--------------------------------------------------------------------------------
-- 1) CABECALHO: quais colunas de NotaFiscal estao NULL nessa nota
--------------------------------------------------------------------------------
DECLARE @sql NVARCHAR(MAX);

SET @sql = N'';
SELECT @sql = @sql +
    N' UNION ALL SELECT ''NotaFiscal'' AS tabela, NULL AS item, ''' + c.name + N''' AS coluna' +
    N' FROM NotaFiscal WHERE CodigoNota = @cn AND ' + QUOTENAME(c.name) + N' IS NULL'
FROM sys.columns c
WHERE c.object_id = OBJECT_ID('NotaFiscal');

SET @sql = STUFF(@sql, 1, 11, N'') + N' ORDER BY coluna';
EXEC sp_executesql @sql, N'@cn INT', @cn = @CodigoNota;

--------------------------------------------------------------------------------
-- 2) ITENS: quais colunas de NotaFiscalItens estao NULL, por item
--------------------------------------------------------------------------------
SET @sql = N'';
SELECT @sql = @sql +
    N' UNION ALL SELECT ''NotaFiscalItens'' AS tabela, CAST(ITEM AS VARCHAR(10)) AS item, ''' + c.name + N''' AS coluna' +
    N' FROM NotaFiscalItens WHERE CodigoNota = @cn AND ' + QUOTENAME(c.name) + N' IS NULL'
FROM sys.columns c
WHERE c.object_id = OBJECT_ID('NotaFiscalItens');

SET @sql = STUFF(@sql, 1, 11, N'') + N' ORDER BY item, coluna';
EXEC sp_executesql @sql, N'@cn INT', @cn = @CodigoNota;

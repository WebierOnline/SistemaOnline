-- Protege contra dado corrompido (valor fora da faixa do novo tipo) antes de cada
-- ALTER. Sem isso, o SQL Server tenta converter a coluna inteira e falha (Msg 8115 -
-- estouro aritmetico) se existir UMA linha so com lixo (ex: codigo de barra gravado
-- no lugar do valor - ja visto em clientes reais). Os limiares abaixo sao exatamente
-- o maximo que decimal(16,2)/decimal(16,3) aceitam - zera so o que ja ia estourar.

-- Desliga AUTO_CREATE_STATISTICS no banco durante este script -
-- confirmado num cliente real que o proprio ALTER TABLE ALTER
-- COLUMN recria a estatistica auto no meio da sua propria execucao
-- (scan interno de validacao do tipo), entao so dropar antes nao
-- resolve (testado 2x, falhou do mesmo jeito). Guarda o estado
-- original pra restaurar no final, mesmo que ja estivesse OFF.
DECLARE @autoStatsEstadoOriginal BIT;
SELECT @autoStatsEstadoOriginal = is_auto_create_stats_on FROM sys.databases WHERE database_id = DB_ID();
DECLARE @sqlAutoStatsOff NVARCHAR(300) = N'ALTER DATABASE ' + QUOTENAME(DB_NAME()) + N' SET AUTO_CREATE_STATISTICS OFF';
EXEC(@sqlAutoStatsOff);

UPDATE pedidos_itens SET PRECO = 0 WHERE PRECO >= 99999999999999.99 OR PRECO <= -99999999999999.99;
UPDATE pedidos_itens SET Desconto = 0 WHERE Desconto >= 99999999999999.99 OR Desconto <= -99999999999999.99;
UPDATE pedidos_itens SET Subtotal = 0 WHERE Subtotal >= 99999999999999.99 OR Subtotal <= -99999999999999.99;
UPDATE pedidos_itens SET Total = 0 WHERE Total >= 99999999999999.99 OR Total <= -99999999999999.99;
UPDATE pedidos_itens SET QUANTIDADE = 0 WHERE QUANTIDADE >= 9999999999999.999 OR QUANTIDADE <= -9999999999999.999;

DECLARE @dropStats_pedidos_itens NVARCHAR(MAX) = N'';
SELECT @dropStats_pedidos_itens = @dropStats_pedidos_itens + 'DROP STATISTICS ' + QUOTENAME(OBJECT_SCHEMA_NAME(s.object_id)) + '.' + QUOTENAME(OBJECT_NAME(s.object_id)) + '.' + QUOTENAME(s.name) + ';'
FROM sys.stats s
JOIN sys.stats_columns sc ON s.object_id = sc.object_id AND s.stats_id = sc.stats_id
JOIN sys.columns c ON sc.object_id = c.object_id AND sc.column_id = c.column_id
WHERE s.object_id = OBJECT_ID('pedidos_itens') AND c.name IN ('PRECO','QUANTIDADE','Desconto','Subtotal','Total');
EXEC sp_executesql @dropStats_pedidos_itens;

ALTER TABLE pedidos_itens ALTER COLUMN PRECO decimal(16,2);
ALTER TABLE pedidos_itens ALTER COLUMN QUANTIDADE decimal(16,3);
ALTER TABLE pedidos_itens ALTER COLUMN Desconto decimal(16,2);
ALTER TABLE pedidos_itens ALTER COLUMN Subtotal decimal(16,2);
ALTER TABLE pedidos_itens ALTER COLUMN Total decimal(16,2);

UPDATE pedidos SET SUBTOTAL = 0 WHERE SUBTOTAL >= 99999999999999.99 OR SUBTOTAL <= -99999999999999.99;
UPDATE pedidos SET TOTAL = 0 WHERE TOTAL >= 99999999999999.99 OR TOTAL <= -99999999999999.99;
UPDATE pedidos SET ENTRADA = 0 WHERE ENTRADA >= 99999999999999.99 OR ENTRADA <= -99999999999999.99;
UPDATE pedidos SET VALOR_DESC = 0 WHERE VALOR_DESC >= 99999999999999.99 OR VALOR_DESC <= -99999999999999.99;
UPDATE pedidos SET VALOR_ACRESCIMO = 0 WHERE VALOR_ACRESCIMO >= 99999999999999.99 OR VALOR_ACRESCIMO <= -99999999999999.99;
UPDATE pedidos SET ValorDescReal = 0 WHERE ValorDescReal >= 99999999999999.99 OR ValorDescReal <= -99999999999999.99;
UPDATE pedidos SET ValorAcrescReal = 0 WHERE ValorAcrescReal >= 99999999999999.99 OR ValorAcrescReal <= -99999999999999.99;

DECLARE @dropStats_pedidos NVARCHAR(MAX) = N'';
SELECT @dropStats_pedidos = @dropStats_pedidos + 'DROP STATISTICS ' + QUOTENAME(OBJECT_SCHEMA_NAME(s.object_id)) + '.' + QUOTENAME(OBJECT_NAME(s.object_id)) + '.' + QUOTENAME(s.name) + ';'
FROM sys.stats s
JOIN sys.stats_columns sc ON s.object_id = sc.object_id AND s.stats_id = sc.stats_id
JOIN sys.columns c ON sc.object_id = c.object_id AND sc.column_id = c.column_id
WHERE s.object_id = OBJECT_ID('pedidos') AND c.name IN ('SUBTOTAL','TOTAL','ENTRADA','VALOR_DESC','VALOR_ACRESCIMO','ValorDescReal','ValorAcrescReal');
EXEC sp_executesql @dropStats_pedidos;

ALTER TABLE pedidos ALTER COLUMN SUBTOTAL decimal(16,2);
ALTER TABLE pedidos ALTER COLUMN TOTAL decimal(16,2);
ALTER TABLE pedidos ALTER COLUMN ENTRADA decimal(16,2);
ALTER TABLE pedidos ALTER COLUMN VALOR_DESC decimal(16,2);
ALTER TABLE pedidos ALTER COLUMN VALOR_ACRESCIMO decimal(16,2);
ALTER TABLE pedidos ALTER COLUMN ValorDescReal decimal(16,2);
ALTER TABLE pedidos ALTER COLUMN ValorAcrescReal decimal(16,2);

UPDATE produtos SET quant_estoque = 0 WHERE quant_estoque >= 9999999999999.999 OR quant_estoque <= -9999999999999.999;
UPDATE produtos SET quant_min = 0 WHERE quant_min >= 9999999999999.999 OR quant_min <= -9999999999999.999;

DECLARE @dropStats_produtos NVARCHAR(MAX) = N'';
SELECT @dropStats_produtos = @dropStats_produtos + 'DROP STATISTICS ' + QUOTENAME(OBJECT_SCHEMA_NAME(s.object_id)) + '.' + QUOTENAME(OBJECT_NAME(s.object_id)) + '.' + QUOTENAME(s.name) + ';'
FROM sys.stats s
JOIN sys.stats_columns sc ON s.object_id = sc.object_id AND s.stats_id = sc.stats_id
JOIN sys.columns c ON sc.object_id = c.object_id AND sc.column_id = c.column_id
WHERE s.object_id = OBJECT_ID('produtos') AND c.name IN ('quant_estoque','quant_min');
EXEC sp_executesql @dropStats_produtos;

ALTER TABLE produtos ALTER COLUMN quant_estoque decimal(16,3);
ALTER TABLE produtos ALTER COLUMN quant_min decimal(16,3);

UPDATE TbNFCe_Itens SET QtdeMov = 0 WHERE QtdeMov >= 9999999999999.999 OR QtdeMov <= -9999999999999.999;
UPDATE TbNFCe_Itens SET desconto = 0 WHERE desconto >= 99999999999999.99 OR desconto <= -99999999999999.99;
UPDATE TbNFCe_Itens SET ValorUnit = 0 WHERE ValorUnit >= 99999999999999.99 OR ValorUnit <= -99999999999999.99;

DECLARE @dropStats_TbNFCe_Itens NVARCHAR(MAX) = N'';
SELECT @dropStats_TbNFCe_Itens = @dropStats_TbNFCe_Itens + 'DROP STATISTICS ' + QUOTENAME(OBJECT_SCHEMA_NAME(s.object_id)) + '.' + QUOTENAME(OBJECT_NAME(s.object_id)) + '.' + QUOTENAME(s.name) + ';'
FROM sys.stats s
JOIN sys.stats_columns sc ON s.object_id = sc.object_id AND s.stats_id = sc.stats_id
JOIN sys.columns c ON sc.object_id = c.object_id AND sc.column_id = c.column_id
WHERE s.object_id = OBJECT_ID('TbNFCe_Itens') AND c.name IN ('QtdeMov','desconto','ValorUnit');
EXEC sp_executesql @dropStats_TbNFCe_Itens;

ALTER TABLE TbNFCe_Itens ALTER COLUMN QtdeMov decimal(16,3);
ALTER TABLE TbNFCe_Itens ALTER COLUMN desconto decimal(16,2);
ALTER TABLE TbNFCe_Itens ALTER COLUMN ValorUnit decimal(16,2);

UPDATE NotaFiscalItens SET ValorTotalBruto = 0 WHERE ValorTotalBruto >= 99999999999999.99 OR ValorTotalBruto <= -99999999999999.99;
UPDATE NotaFiscalItens SET ValorUnitarioComercializacao = 0 WHERE ValorUnitarioComercializacao >= 99999999999999.99 OR ValorUnitarioComercializacao <= -99999999999999.99;
UPDATE NotaFiscalItens SET vBC = 0 WHERE vBC >= 99999999999999.99 OR vBC <= -99999999999999.99;
UPDATE NotaFiscalItens SET Desconto = 0 WHERE Desconto >= 99999999999999.99 OR Desconto <= -99999999999999.99;
UPDATE NotaFiscalItens SET ValorDesconto = 0 WHERE ValorDesconto >= 99999999999999.99 OR ValorDesconto <= -99999999999999.99;
UPDATE NotaFiscalItens SET QuantidadeComercial = 0 WHERE QuantidadeComercial >= 9999999999999.999 OR QuantidadeComercial <= -9999999999999.999;

DECLARE @dropStats_NotaFiscalItens NVARCHAR(MAX) = N'';
SELECT @dropStats_NotaFiscalItens = @dropStats_NotaFiscalItens + 'DROP STATISTICS ' + QUOTENAME(OBJECT_SCHEMA_NAME(s.object_id)) + '.' + QUOTENAME(OBJECT_NAME(s.object_id)) + '.' + QUOTENAME(s.name) + ';'
FROM sys.stats s
JOIN sys.stats_columns sc ON s.object_id = sc.object_id AND s.stats_id = sc.stats_id
JOIN sys.columns c ON sc.object_id = c.object_id AND sc.column_id = c.column_id
WHERE s.object_id = OBJECT_ID('NotaFiscalItens') AND c.name IN ('ValorTotalBruto','ValorUnitarioComercializacao','vBC','Desconto','ValorDesconto','QuantidadeComercial');
EXEC sp_executesql @dropStats_NotaFiscalItens;

ALTER TABLE NotaFiscalItens ALTER COLUMN ValorTotalBruto decimal(16,2);
ALTER TABLE NotaFiscalItens ALTER COLUMN ValorUnitarioComercializacao decimal(16,2);
ALTER TABLE NotaFiscalItens ALTER COLUMN vBC decimal(16,2);
ALTER TABLE NotaFiscalItens ALTER COLUMN Desconto decimal(16,2);
ALTER TABLE NotaFiscalItens ALTER COLUMN ValorDesconto decimal(16,2);
ALTER TABLE NotaFiscalItens ALTER COLUMN QuantidadeComercial decimal(16,3);

UPDATE NotaFiscal SET ValorProdutos = 0 WHERE ValorProdutos >= 99999999999999.99 OR ValorProdutos <= -99999999999999.99;
UPDATE NotaFiscal SET ValorNota = 0 WHERE ValorNota >= 99999999999999.99 OR ValorNota <= -99999999999999.99;
UPDATE NotaFiscal SET ValorOriginalFatura = 0 WHERE ValorOriginalFatura >= 99999999999999.99 OR ValorOriginalFatura <= -99999999999999.99;
UPDATE NotaFiscal SET ValorLiquidoFatura = 0 WHERE ValorLiquidoFatura >= 99999999999999.99 OR ValorLiquidoFatura <= -99999999999999.99;

DECLARE @dropStats_NotaFiscal NVARCHAR(MAX) = N'';
SELECT @dropStats_NotaFiscal = @dropStats_NotaFiscal + 'DROP STATISTICS ' + QUOTENAME(OBJECT_SCHEMA_NAME(s.object_id)) + '.' + QUOTENAME(OBJECT_NAME(s.object_id)) + '.' + QUOTENAME(s.name) + ';'
FROM sys.stats s
JOIN sys.stats_columns sc ON s.object_id = sc.object_id AND s.stats_id = sc.stats_id
JOIN sys.columns c ON sc.object_id = c.object_id AND sc.column_id = c.column_id
WHERE s.object_id = OBJECT_ID('NotaFiscal') AND c.name IN ('ValorProdutos','ValorNota','ValorOriginalFatura','ValorLiquidoFatura');
EXEC sp_executesql @dropStats_NotaFiscal;

ALTER TABLE NotaFiscal ALTER COLUMN ValorProdutos decimal(16,2);
ALTER TABLE NotaFiscal ALTER COLUMN ValorNota decimal(16,2);
ALTER TABLE NotaFiscal ALTER COLUMN ValorOriginalFatura decimal(16,2);
ALTER TABLE NotaFiscal ALTER COLUMN ValorLiquidoFatura decimal(16,2);

UPDATE parcelas SET JUROS = 0 WHERE JUROS >= 99999999999999.99 OR JUROS <= -99999999999999.99;
UPDATE parcelas SET MULTA = 0 WHERE MULTA >= 99999999999999.99 OR MULTA <= -99999999999999.99;
UPDATE parcelas SET VALOR = 0 WHERE VALOR >= 99999999999999.99 OR VALOR <= -99999999999999.99;
UPDATE parcelas SET VALOR_FINAL = 0 WHERE VALOR_FINAL >= 99999999999999.99 OR VALOR_FINAL <= -99999999999999.99;
UPDATE parcelas SET DESCONTO = 0 WHERE DESCONTO >= 99999999999999.99 OR DESCONTO <= -99999999999999.99;

DECLARE @dropStats_parcelas NVARCHAR(MAX) = N'';
SELECT @dropStats_parcelas = @dropStats_parcelas + 'DROP STATISTICS ' + QUOTENAME(OBJECT_SCHEMA_NAME(s.object_id)) + '.' + QUOTENAME(OBJECT_NAME(s.object_id)) + '.' + QUOTENAME(s.name) + ';'
FROM sys.stats s
JOIN sys.stats_columns sc ON s.object_id = sc.object_id AND s.stats_id = sc.stats_id
JOIN sys.columns c ON sc.object_id = c.object_id AND sc.column_id = c.column_id
WHERE s.object_id = OBJECT_ID('parcelas') AND c.name IN ('JUROS','MULTA','VALOR','VALOR_FINAL','DESCONTO');
EXEC sp_executesql @dropStats_parcelas;

ALTER TABLE parcelas ALTER COLUMN JUROS decimal(16,2);
ALTER TABLE parcelas ALTER COLUMN MULTA decimal(16,2);
ALTER TABLE parcelas ALTER COLUMN VALOR decimal(16,2);
ALTER TABLE parcelas ALTER COLUMN VALOR_FINAL decimal(16,2);
ALTER TABLE parcelas ALTER COLUMN DESCONTO decimal(16,2);

UPDATE Produtos_Precos SET CUSTO = 0 WHERE CUSTO >= 99999999999999.99 OR CUSTO <= -99999999999999.99;
UPDATE Produtos_Precos SET VALOR_VV = 0 WHERE VALOR_VV >= 99999999999999.99 OR VALOR_VV <= -99999999999999.99;
UPDATE Produtos_Precos SET VALOR_VP = 0 WHERE VALOR_VP >= 99999999999999.99 OR VALOR_VP <= -99999999999999.99;
UPDATE Produtos_Precos SET VALOR_AV = 0 WHERE VALOR_AV >= 99999999999999.99 OR VALOR_AV <= -99999999999999.99;
UPDATE Produtos_Precos SET VALOR_AP = 0 WHERE VALOR_AP >= 99999999999999.99 OR VALOR_AP <= -99999999999999.99;
UPDATE Produtos_Precos SET MARGEM_VV = 0 WHERE MARGEM_VV >= 99999999999999.99 OR MARGEM_VV <= -99999999999999.99;
UPDATE Produtos_Precos SET MARGEM_VP = 0 WHERE MARGEM_VP >= 99999999999999.99 OR MARGEM_VP <= -99999999999999.99;
UPDATE Produtos_Precos SET MARGEM_AV = 0 WHERE MARGEM_AV >= 99999999999999.99 OR MARGEM_AV <= -99999999999999.99;
UPDATE Produtos_Precos SET MARGEM_AP = 0 WHERE MARGEM_AP >= 99999999999999.99 OR MARGEM_AP <= -99999999999999.99;

DECLARE @dropStats_Produtos_Precos NVARCHAR(MAX) = N'';
SELECT @dropStats_Produtos_Precos = @dropStats_Produtos_Precos + 'DROP STATISTICS ' + QUOTENAME(OBJECT_SCHEMA_NAME(s.object_id)) + '.' + QUOTENAME(OBJECT_NAME(s.object_id)) + '.' + QUOTENAME(s.name) + ';'
FROM sys.stats s
JOIN sys.stats_columns sc ON s.object_id = sc.object_id AND s.stats_id = sc.stats_id
JOIN sys.columns c ON sc.object_id = c.object_id AND sc.column_id = c.column_id
WHERE s.object_id = OBJECT_ID('Produtos_Precos') AND c.name IN ('CUSTO','VALOR_VV','VALOR_VP','VALOR_AV','VALOR_AP','MARGEM_VV','MARGEM_VP','MARGEM_AV','MARGEM_AP');
EXEC sp_executesql @dropStats_Produtos_Precos;

ALTER TABLE Produtos_Precos ALTER COLUMN CUSTO decimal(16,2);
ALTER TABLE Produtos_Precos ALTER COLUMN VALOR_VV decimal(16,2);
ALTER TABLE Produtos_Precos ALTER COLUMN VALOR_VP decimal(16,2);
ALTER TABLE Produtos_Precos ALTER COLUMN VALOR_AV decimal(16,2);
ALTER TABLE Produtos_Precos ALTER COLUMN VALOR_AP decimal(16,2);
ALTER TABLE Produtos_Precos ALTER COLUMN MARGEM_VV decimal(16,2);
ALTER TABLE Produtos_Precos ALTER COLUMN MARGEM_VP decimal(16,2);
ALTER TABLE Produtos_Precos ALTER COLUMN MARGEM_AV decimal(16,2);
ALTER TABLE Produtos_Precos ALTER COLUMN MARGEM_AP decimal(16,2);

IF OBJECT_ID('a_receber_itens', 'U') IS NOT NULL
BEGIN
    EXEC('UPDATE a_receber_itens SET preco = 0 WHERE preco >= 99999999999999.99 OR preco <= -99999999999999.99');

    DECLARE @dropStats_areceber NVARCHAR(MAX) = N'';
    SELECT @dropStats_areceber = ISNULL(@dropStats_areceber,N'') + 'DROP STATISTICS ' + QUOTENAME(OBJECT_SCHEMA_NAME(s.object_id)) + '.' + QUOTENAME(OBJECT_NAME(s.object_id)) + '.' + QUOTENAME(s.name) + ';'
    FROM sys.stats s
    JOIN sys.stats_columns sc ON s.object_id = sc.object_id AND s.stats_id = sc.stats_id
    JOIN sys.columns c ON sc.object_id = c.object_id AND sc.column_id = c.column_id
    WHERE s.object_id = OBJECT_ID('a_receber_itens') AND c.name = 'preco';
    EXEC sp_executesql @dropStats_areceber;

    EXEC('ALTER TABLE a_receber_itens ALTER COLUMN preco decimal(16,2)');
END

IF @autoStatsEstadoOriginal = 1
BEGIN
   DECLARE @sqlAutoStatsOn NVARCHAR(300) = N'ALTER DATABASE ' + QUOTENAME(DB_NAME()) + N' SET AUTO_CREATE_STATISTICS ON';
   EXEC(@sqlAutoStatsOn);
END

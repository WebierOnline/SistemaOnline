-- Corrige OS.SUBTOTAL das OS ja gravadas com o subtotal errado.
-- O sistema (Atualizar_Dados_OS) gravava em OS.SUBTOTAL o total JA COM DESCONTO dos itens; o correto e o
-- subtotal ANTES dos descontos = soma dos subtotais dos servicos da OS + soma dos subtotais dos itens do pedido
-- (o mesmo valor do campo Subtotal da tela). So o SUBTOTAL e alterado: TOTAL, VALOR_DESC e ValorDescReal
-- nao sao tocados.
-- OS sem nenhum item (soma = 0) ficam como estao. Tabelas de servicos que nao existirem no cliente sao ignoradas.
-- Idempotente: so atualiza as OS cujo SUBTOTAL difere da soma dos itens (rodar de novo nao altera nada).
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

IF OBJECT_ID('dbo.OS', 'U') IS NOT NULL
   AND COL_LENGTH('dbo.OS', 'SUBTOTAL') IS NOT NULL
   AND COL_LENGTH('dbo.OS', 'COD_PEDIDO') IS NOT NULL
BEGIN
    DECLARE @sql nvarchar(max)

    -- soma dos subtotais dos itens de cada OS (servicos das tabelas existentes + itens do pedido)
    SET @sql = N'
    IF OBJECT_ID(''tempdb..#os_sub'') IS NOT NULL DROP TABLE #os_sub;
    SELECT o.COD_OS, CAST(0 AS decimal(18,2)) AS Soma INTO #os_sub FROM dbo.OS o;'

    IF OBJECT_ID('dbo.OS_Servicos_Auto', 'U') IS NOT NULL
        SET @sql = @sql + N'
    UPDATE t SET Soma = t.Soma + s.v FROM #os_sub t
    INNER JOIN (SELECT cod_os, SUM(ISNULL(subtotal, 0)) AS v FROM dbo.OS_Servicos_Auto GROUP BY cod_os) s ON s.cod_os = t.COD_OS;'

    IF OBJECT_ID('dbo.OS_Servicos_Recapadora', 'U') IS NOT NULL
        SET @sql = @sql + N'
    UPDATE t SET Soma = t.Soma + s.v FROM #os_sub t
    INNER JOIN (SELECT cod_os, SUM(ISNULL(subtotal, 0)) AS v FROM dbo.OS_Servicos_Recapadora GROUP BY cod_os) s ON s.cod_os = t.COD_OS;'

    IF OBJECT_ID('dbo.OS_Servicos_Comunicacao', 'U') IS NOT NULL
        SET @sql = @sql + N'
    UPDATE t SET Soma = t.Soma + s.v FROM #os_sub t
    INNER JOIN (SELECT cod_os, SUM(ISNULL(subtotal, 0)) AS v FROM dbo.OS_Servicos_Comunicacao GROUP BY cod_os) s ON s.cod_os = t.COD_OS;'

    IF OBJECT_ID('dbo.pedidos_itens', 'U') IS NOT NULL
        SET @sql = @sql + N'
    UPDATE t SET Soma = t.Soma + p.v FROM #os_sub t
    INNER JOIN dbo.OS o ON o.COD_OS = t.COD_OS
    INNER JOIN (SELECT cod_pedido, SUM(ISNULL(subtotal, 0)) AS v FROM dbo.pedidos_itens GROUP BY cod_pedido) p ON p.cod_pedido = o.COD_PEDIDO;'

    SET @sql = @sql + N'
    UPDATE o SET o.SUBTOTAL = t.Soma
    FROM dbo.OS o INNER JOIN #os_sub t ON t.COD_OS = o.COD_OS
    WHERE t.Soma > 0 AND ISNULL(o.SUBTOTAL, -1) <> t.Soma;
    PRINT ''OS com SUBTOTAL corrigido: '' + CAST(@@ROWCOUNT AS varchar(20));
    DROP TABLE #os_sub;'

    EXEC (@sql)
END
GO

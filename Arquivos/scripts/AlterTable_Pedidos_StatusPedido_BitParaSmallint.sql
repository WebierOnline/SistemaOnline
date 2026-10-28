-- pedidos.status_pedido esta como BIT (so aceita 0/1), mas o codigo sempre tratou como inteiro:
--   0 = aberto, 1 = fechado, -1 = pausado (PDV.frm cmdAvanVendaPausar_Click / Estonar.frm cboStatus "PAUSADO"), 3 = referenciado em Estonar (filtro "incompleto")
-- Como BIT so guarda 0/1, gravar -1 e silenciosamente convertido para 1 pelo SQL Server (qualquer valor <> 0 vira 1).
-- Resultado: pausar uma venda marcava o pedido como "fechado" (status=1) em vez de "pausado", e toda busca
-- por status_pedido = -1 (Reiniciar Venda no PDV, filtro PAUSADO no Estonar) nunca encontrava nada.
IF EXISTS (
    SELECT 1 FROM sys.columns c
    JOIN sys.types t ON c.system_type_id = t.system_type_id AND c.user_type_id = t.user_type_id
    WHERE c.object_id = OBJECT_ID('pedidos') AND c.name = 'status_pedido' AND t.name = 'bit'
)
BEGIN
    -- estatistica auto-criada (_WA_Sys_...) bloqueava o ALTER (Msg 5074/4922,
    -- confirmado em cliente real) - dropa direto por object_id+coluna (o filtro
    -- auto_created=1 nao bate de forma confiavel) e desliga/religa
    -- AUTO_CREATE_STATISTICS no banco em volta do ALTER
    DECLARE @autoStatsEstadoOriginal090 BIT;
    SELECT @autoStatsEstadoOriginal090 = is_auto_create_stats_on FROM sys.databases WHERE database_id = DB_ID();
    DECLARE @sqlAutoStatsOff090 NVARCHAR(300) = N'ALTER DATABASE ' + QUOTENAME(DB_NAME()) + N' SET AUTO_CREATE_STATISTICS OFF';
    EXEC(@sqlAutoStatsOff090);

    DECLARE @dropStats_pedidos_status NVARCHAR(MAX) = N'';
    SELECT @dropStats_pedidos_status = @dropStats_pedidos_status + 'DROP STATISTICS ' + QUOTENAME(OBJECT_SCHEMA_NAME(s.object_id)) + '.' + QUOTENAME(OBJECT_NAME(s.object_id)) + '.' + QUOTENAME(s.name) + ';'
    FROM sys.stats s
    JOIN sys.stats_columns sc ON s.object_id = sc.object_id AND s.stats_id = sc.stats_id
    JOIN sys.columns c ON sc.object_id = c.object_id AND sc.column_id = c.column_id
    WHERE s.object_id = OBJECT_ID('pedidos') AND c.name = 'status_pedido';
    EXEC sp_executesql @dropStats_pedidos_status;

    ALTER TABLE pedidos ALTER COLUMN status_pedido SMALLINT NULL;

    IF @autoStatsEstadoOriginal090 = 1
    BEGIN
        DECLARE @sqlAutoStatsOn090 NVARCHAR(300) = N'ALTER DATABASE ' + QUOTENAME(DB_NAME()) + N' SET AUTO_CREATE_STATISTICS ON';
        EXEC(@sqlAutoStatsOn090);
    END
END

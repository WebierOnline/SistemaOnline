-- TbNotasDestinadasManifestacao: valores padrao nas colunas obrigatorias.
-- O CREATE TABLE do script 133 nao tem DEFAULT, entao o INSERT da tela Notas_Destinadas_Consulta (que so informa
-- algumas colunas) falhava com "nao e possivel inserir NULL". Idempotente: so cria o DEFAULT onde a coluna existe e
-- ainda nao tem nenhum. Cria tambem o indice por ChaveAcesso (consultado a cada nota recebida pelo monitor).
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET NOCOUNT ON
GO

IF OBJECT_ID('dbo.TbNotasDestinadasManifestacao', 'U') IS NOT NULL
BEGIN
    DECLARE @padroes TABLE (Coluna sysname, Valor nvarchar(20))
    INSERT INTO @padroes (Coluna, Valor) VALUES
        ('CodigoConsulta', '(0)'), ('TipoRetornoWS', ''''''), ('TipoNotaWS', '(0)'), ('CNPJFornecedor', ''''''),
        ('RazaoSocial', ''''''), ('ChaveAcesso', ''''''), ('ValorNota', '(0)'), ('SituacaoNota', '(0)'),
        ('SituacaoConfirmacao', '(0)'), ('nSeqEvento', '(0)'), ('TextoCorrecao', ''''''), ('TipoEvento', ''''''),
        ('Justificativa', ''''''), ('NumeroProtocolo', '(0)'), ('DataHoraProcotolo', ''''''), ('Enviada', '(0)'),
        ('Enviar', '(0)'), ('XMLAutorizado', ''''''), ('ImportadoXML', '(0)'), ('IdFilial', '(1)'),
        ('TipoEventoCodigo', ''''''), ('EventoProtocolo', '(0)'), ('EventoDataHora', '''''')

    DECLARE @col sysname, @val nvarchar(20), @sql nvarchar(400)
    DECLARE c CURSOR LOCAL FAST_FORWARD FOR
        SELECT p.Coluna, p.Valor
        FROM @padroes p
        JOIN sys.columns sc ON sc.object_id = OBJECT_ID('dbo.TbNotasDestinadasManifestacao') AND sc.name = p.Coluna
        WHERE sc.default_object_id = 0
    OPEN c
    FETCH NEXT FROM c INTO @col, @val
    WHILE @@FETCH_STATUS = 0
    BEGIN
        SET @sql = N'ALTER TABLE dbo.TbNotasDestinadasManifestacao ADD CONSTRAINT ' + QUOTENAME('DF_TbNotasDestinadas_' + @col) +
                   N' DEFAULT (' + @val + N') FOR ' + QUOTENAME(@col)
        EXEC (@sql)
        FETCH NEXT FROM c INTO @col, @val
    END
    CLOSE c
    DEALLOCATE c
END
GO

IF OBJECT_ID('dbo.TbNotasDestinadasManifestacao', 'U') IS NOT NULL
   AND NOT EXISTS (SELECT 1 FROM sys.indexes WHERE object_id = OBJECT_ID('dbo.TbNotasDestinadasManifestacao') AND name = 'IX_TbNotasDestinadas_ChaveAcesso')
    CREATE NONCLUSTERED INDEX IX_TbNotasDestinadas_ChaveAcesso ON dbo.TbNotasDestinadasManifestacao (ChaveAcesso)
GO

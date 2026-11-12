-- Importacao automatica das NF-e destinadas pelo monitor (/monitornfe) na Entrada de Estoque.
-- EntradaEstoque.OrigemLancamento: 'MANUAL' (tela) ou 'MONITOR' (importada sozinha). Linhas antigas ficam NULL = manual.
-- TbNotasDestinadasManifestacao: ImportacaoStatus (0 pendente, 1 importada, 2 erro, 3 nao e compra / nao importar),
-- ImportacaoMotivo, ImportacaoData e CodigoNotaEntrada (EntradaEstoque.CodigoNota) - o monitor so olha as pendentes.
-- configuracao: Importar_XML (1 = importa, padrao) e IMPORTAR_XML_DESDE (data a partir da qual importa; o monitor grava
-- a data de hoje na primeira execucao, para nao importar o historico de 3 meses que a SEFAZ devolve na primeira consulta).
-- View NotasDestinadas: situacao "Erro Importação" e coluna ImportacaoMotivo. Idempotente.
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

IF COL_LENGTH('dbo.EntradaEstoque', 'OrigemLancamento') IS NULL
    ALTER TABLE dbo.EntradaEstoque ADD OrigemLancamento varchar(10) NULL
GO

IF OBJECT_ID('dbo.TbNotasDestinadasManifestacao', 'U') IS NOT NULL
BEGIN
    IF COL_LENGTH('dbo.TbNotasDestinadasManifestacao', 'ImportacaoStatus') IS NULL
        ALTER TABLE dbo.TbNotasDestinadasManifestacao ADD ImportacaoStatus tinyint NOT NULL
            CONSTRAINT DF_TbNotasDestinadas_ImportacaoStatus DEFAULT (0)
    IF COL_LENGTH('dbo.TbNotasDestinadasManifestacao', 'ImportacaoMotivo') IS NULL
        ALTER TABLE dbo.TbNotasDestinadasManifestacao ADD ImportacaoMotivo varchar(250) NOT NULL
            CONSTRAINT DF_TbNotasDestinadas_ImportacaoMotivo DEFAULT ('')
    IF COL_LENGTH('dbo.TbNotasDestinadasManifestacao', 'ImportacaoData') IS NULL
        ALTER TABLE dbo.TbNotasDestinadasManifestacao ADD ImportacaoData datetime NULL
    IF COL_LENGTH('dbo.TbNotasDestinadasManifestacao', 'CodigoNotaEntrada') IS NULL
        ALTER TABLE dbo.TbNotasDestinadasManifestacao ADD CodigoNotaEntrada int NOT NULL
            CONSTRAINT DF_TbNotasDestinadas_CodigoNotaEntrada DEFAULT (0)
END
GO

-- notas que ja estao na Entrada de Estoque nao sao importadas de novo
IF OBJECT_ID('dbo.TbNotasDestinadasManifestacao', 'U') IS NOT NULL
    UPDATE d SET ImportacaoStatus = 1, CodigoNotaEntrada = e.CodigoNota
    FROM dbo.TbNotasDestinadasManifestacao d
    CROSS APPLY (SELECT TOP 1 CodigoNota FROM dbo.EntradaEstoque WHERE ChavedeAcesso = d.ChaveAcesso ORDER BY CodigoNota DESC) e
    WHERE d.ImportacaoStatus = 0
GO

IF NOT EXISTS (SELECT 1 FROM configuracao WHERE config_nome = 'Importar_XML')
    INSERT INTO configuracao (config_nome, config_valor) VALUES ('Importar_XML', '1')
IF NOT EXISTS (SELECT 1 FROM configuracao WHERE config_nome = 'IMPORTAR_XML_DESDE')
    INSERT INTO configuracao (config_nome, config_valor) VALUES ('IMPORTAR_XML_DESDE', '')
GO

IF OBJECT_ID('dbo.NotasDestinadas') IS NULL
    EXEC('CREATE VIEW dbo.NotasDestinadas AS SELECT 1 AS Placeholder')
GO

ALTER VIEW dbo.NotasDestinadas AS
SELECT (CASE
          WHEN ISNULL(TipoEventoCodigo, '') = '110111' THEN 'Cancelada'
          WHEN ISNULL(TipoEventoCodigo, '') = '210220' OR ISNULL(TipoEvento, '') LIKE '2 - %' THEN 'Desconhecida'
          WHEN ISNULL(TipoEventoCodigo, '') = '210240' OR ISNULL(TipoEvento, '') LIKE '3 - %' THEN 'Não Realizada'
          WHEN dbo.SituacaoNF(TbNotasDestinadasManifestacao.ChaveAcesso) = 1 THEN (CASE WHEN Enviada = 0 THEN 'Importada/Sem Manifesto' ELSE 'Importada' END)
          WHEN EXISTS (SELECT 1 FROM EntradaEstoque e WHERE e.ChavedeAcesso = TbNotasDestinadasManifestacao.ChaveAcesso AND ISNULL(e.Enviada, 0) = 0) THEN 'Em Lançamento'
          WHEN ImportacaoStatus = 2 THEN 'Erro Importação'
          WHEN ISNULL(TipoEventoCodigo, '') = '210200' OR ISNULL(TipoEvento, '') LIKE '0 - %' THEN 'Confirmada'
          WHEN Enviada = 1 THEN 'Manifestada'
          WHEN ISNULL(TipoEventoCodigo, '') = '210210' OR ISNULL(TipoEvento, '') LIKE '1 - %' THEN 'Aguardando XML'
          ELSE 'Sem Manifestar'
        END) AS SituacaoNF,
TipoRetornoWS, (CASE WHEN TipoEvento <> '' THEN (CASE WHEN ISNUMERIC(LEFT(TipoEvento, 1)) = 1 THEN SUBSTRING(TipoEvento, 5, 100) ELSE TipoEvento END) ELSE '' END) AS TipoEvento, TipoNotaWS, Data, CNPJFornecedor, RazaoSocial, ChaveAcesso, DataEmissao, ValorNota,
SituacaoNota, dbo.SituacaoNF(TbNotasDestinadasManifestacao.ChaveAcesso) As ImportadoXML,
Enviar, Enviada, CodigoConsulta, (CASE WHEN TipoEvento <> '' THEN LEFT(TipoEvento, 1) ELSE '' END) AS TipoEventoID, IdNFConsulta, IdFilial, TipoEventoCodigo, EventoProtocolo, EventoDataHora,
ImportacaoStatus, ImportacaoMotivo, CodigoNotaEntrada
FROM TbNotasDestinadasManifestacao
GO

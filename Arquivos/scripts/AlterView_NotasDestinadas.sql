/***************************
PROJETO...: Online.info
ANALISTA..: Leandro Santos
DESCRICAO.: View de Nota Destinadas para CNPJ da Empresa
DATA......: 2022-03-28
ALTERACAO.: 
***************************/
-- Idempotente: cria um placeholder se o objeto ainda nao existir e sempre roda o ALTER com a definicao final.
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF OBJECT_ID('dbo.NotasDestinadas') IS NULL
    EXEC('CREATE VIEW dbo.NotasDestinadas AS SELECT 1 AS Placeholder')
GO

ALTER VIEW dbo.NotasDestinadas AS
SELECT (CASE WHEN dbo.SituacaoNF(TbNotasDestinadasManifestacao.ChaveAcesso) = 1 AND TipoEventoCodigo <> '110111' THEN (CASE WHEN Enviada = 0 THEN 'Importada/Sem Manifesto' ELSE 'Importada' END) ELSE (CASE WHEN TipoEventoCodigo = '110111' THEN 'Cancelada' ELSE (CASE WHEN Enviada = 1 THEN 'Manifestada' ELSE 'Sem Manifestar' END) END) END) AS SituacaoNF,
TipoRetornoWS, (CASE WHEN TipoEvento <> '' THEN (CASE WHEN ISNUMERIC(LEFT(TipoEvento, 1)) = 1 THEN SUBSTRING(TipoEvento, 5, 100) ELSE TipoEvento END) ELSE '' END) AS TipoEvento, TipoNotaWS, Data, CNPJFornecedor, RazaoSocial, ChaveAcesso, DataEmissao, ValorNota,
SituacaoNota, dbo.SituacaoNF(TbNotasDestinadasManifestacao.ChaveAcesso) As ImportadoXML,
Enviar, Enviada, CodigoConsulta, (CASE WHEN TipoEvento <> '' THEN LEFT(TipoEvento, 1) ELSE '' END) AS TipoEventoID, IdNFConsulta, IdFilial, TipoEventoCodigo, EventoProtocolo, EventoDataHora
FROM TbNotasDestinadasManifestacao
GO

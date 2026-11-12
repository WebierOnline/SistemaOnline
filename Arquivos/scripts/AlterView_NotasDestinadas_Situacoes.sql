-- View NotasDestinadas: situacao separando a Ciencia (provisoria) dos eventos definitivos da manifestacao.
-- Antes toda nota com XML virava "Manifestada" e a tela nao deixava mais Confirmar/Desconhecer/registrar Operacao nao
-- Realizada. Situacoes: Cancelada, Desconhecida, Nao Realizada, Importada (/Sem Manifesto), Em Lancamento (entrada de estoque ainda em digitacao), Confirmada, Manifestada
-- (Ciencia + XML disponivel), Aguardando XML (Ciencia enviada, XML ainda nao chegou), Sem Manifestar.
-- O evento vem do texto gravado pela tela/monitor ("0 - ...".."3 - ...") ou do tpEvento recebido na distribuicao.
-- Idempotente (ALTER sobre a view criada pelo script 146).
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
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
          WHEN ISNULL(TipoEventoCodigo, '') = '210200' OR ISNULL(TipoEvento, '') LIKE '0 - %' THEN 'Confirmada'
          WHEN Enviada = 1 THEN 'Manifestada'
          WHEN ISNULL(TipoEventoCodigo, '') = '210210' OR ISNULL(TipoEvento, '') LIKE '1 - %' THEN 'Aguardando XML'
          ELSE 'Sem Manifestar'
        END) AS SituacaoNF,
TipoRetornoWS, (CASE WHEN TipoEvento <> '' THEN (CASE WHEN ISNUMERIC(LEFT(TipoEvento, 1)) = 1 THEN SUBSTRING(TipoEvento, 5, 100) ELSE TipoEvento END) ELSE '' END) AS TipoEvento, TipoNotaWS, Data, CNPJFornecedor, RazaoSocial, ChaveAcesso, DataEmissao, ValorNota,
SituacaoNota, dbo.SituacaoNF(TbNotasDestinadasManifestacao.ChaveAcesso) As ImportadoXML,
Enviar, Enviada, CodigoConsulta, (CASE WHEN TipoEvento <> '' THEN LEFT(TipoEvento, 1) ELSE '' END) AS TipoEventoID, IdNFConsulta, IdFilial, TipoEventoCodigo, EventoProtocolo, EventoDataHora
FROM TbNotasDestinadasManifestacao
GO

/***************************
PROJETO...: SISC
ANALISTA..: Leandro
DESCRICAO.: Procedure de Processamento de NF Serviço Comunicação SPED
DATA......: 2025-05-07
ALTERACAO.: 
***************************/
-- Idempotente: cria um placeholder se o objeto ainda nao existir e sempre roda o ALTER com a definicao final.
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF OBJECT_ID('dbo.SPEDNFServComunicacao') IS NULL
    EXEC('CREATE PROCEDURE dbo.SPEDNFServComunicacao AS SELECT 1')
GO

ALTER PROCEDURE dbo.SPEDNFServComunicacao(@filial INT, @dataini DATETIME, @datafim DATETIME) AS
DECLARE @msg VARCHAR(200)

--Desabilita a contagem recuperacao/alteracao de dados
SET NOCOUNT ON

--Valor padrão
SET @msg = ''

--Opa, tem mensagem, entao retorna em formato de erro  
IF (@msg <> '')
 BEGIN
	RAISERROR(@msg, 16, 1)
	RETURN
 END

IF NOT (object_id('vSPEDNFServComunicacao') is null)
BEGIN
    DROP TABLE vSPEDNFServComunicacao
END

/*
SELECT * FROM TbNFServComunicacao
*/

SELECT IDLancamento, IDParticipante, Modelo, IndicadorTipoOperacao, IndicadorEmitenteDocFiscal, Situacao, Serie, Numero, DataEmissao, DataEntradaReceb, Valor, Desconto, ValorServico, ValorServicoNT, ValorTerceiros, ValorOutras, CFOP, ICMSCST, BaseICMS, AliquotaICMS, ValorICMS, ValorPIS, ValorCOFINS, TipoAssinante, CodigoIBGE
INTO vSPEDNFServComunicacao
FROM
(
SELECT 0 AS IDLancamento, '' AS IDParticipante, '' AS Modelo, 0 AS IndicadorTipoOperacao, 0 AS IndicadorEmitenteDocFiscal, '' As Situacao, '' AS Serie, '' As Numero,
       '' AS DataEmissao, '' AS DataEntradaReceb, 0 AS Valor, 0 As Desconto, 0 AS ValorServico, 0 AS ValorServicoNT, 0 AS ValorTerceiros, 0 AS ValorOutras, 0 AS CFOP, 
	   '' AS ICMSCST, 0 AS BaseICMS, 0 AS AliquotaICMS, 0 AS ValorICMS, 0 AS ValorPIS, 0 AS ValorCOFINS, 0 AS TipoAssinante,	   
	   0 AS CodigoIBGE
FROM empresa
WHERE 1 = 2
) tb

/*
SELECT IdNFServ AS IDLancamento, TbFornecedores.CPF_CNPJ AS IDParticipante, ModeloNF AS Modelo, IndicadorTipoOperacao, IndicadorEmitenteDocFiscal, SituacaoNota As Situacao, (CASE WHEN ISNUMERIC(Serie) = 0 THEN '000' ELSE dbo.LPAD(Serie, 3, '0') END) AS Serie, NumeNota As Numero,
       DataEmissao, DataEntrada AS DataEntradaReceb, ValorNota AS Valor, ValorDesconto As Desconto, ValorServico, ValorServicoNT, ValorTerceiros, ValorOutras, (CASE WHEN CFOP <> 0 THEN (CASE WHEN CFOP > 4000 THEN CFOP - 4000 ELSE CFOP END) ELSE 0 END) AS CFOP, 
	   ICMSCST, BaseICMS, Aliq_ICMS AS AliquotaICMS, ValorICMS, ValorPIS, ValorCOFINS, TipoAssinante,	   
	   TbFiliais.CodigoIBGE
FROM TbNFServComunicacao
     INNER JOIN TbFornecedores ON TbNFServComunicacao.IDFornecedor = TbFornecedores.IDFornecedor
	 INNER JOIN TbCidades ON TbFornecedores.CidadeFornecedor = TbCidades.NomeCidade AND TbFornecedores.UFFornecedor = TbCidades.UF
	 INNER JOIN TbFiliais ON TbNFServComunicacao.IdFilial = TbFiliais.IdFilial
WHERE Fechado <> 0 AND ModeloNF in ('21', '22') AND DataEntrada BETWEEN @dataini AND @datafim AND TbNFServComunicacao.IdFilial = @filial) tb
*/
GO

/***************************
PROJETO...: SISC
ANALISTA..: Leandro
DESCRICAO.: Procedure de Processamento de Conhecimentos SPED
DATA......: 2025-04-04
ALTERACAO.: 
***************************/
-- Idempotente: cria um placeholder se o objeto ainda nao existir e sempre roda o ALTER com a definicao final.
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF OBJECT_ID('dbo.SPEDConhecimento') IS NULL
    EXEC('CREATE PROCEDURE dbo.SPEDConhecimento AS SELECT 1')
GO

ALTER PROCEDURE dbo.SPEDConhecimento(@filial INT, @dataini DATETIME, @datafim DATETIME) AS
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

IF NOT (object_id('vSPEDConhecimento') is null)
BEGIN
    DROP TABLE vSPEDConhecimento
END

SELECT IDLancamento, IDParticipante, Modelo, Situacao, Serie, Numero, ChaveCTe, DataEmissao, DataEntradaReceb, ChaveCTeRef, Valor, Desconto, Frete, ValorServico, CFOP, ICMSCST, BaseICMS, AliquotaICMS, ValorICMS, CodMunOrigem, CodMunDestino
INTO vSPEDConhecimento
FROM
(
SELECT 0 AS IDLancamento, '' AS IDParticipante, '' AS Modelo, '00' As Situacao, '' AS Serie, '' As Numero,
       '' AS ChaveCTe, '' AS DataEmissao, '' AS DataEntradaReceb, '' AS ChaveCTeRef, 0 AS Valor, 0 As Desconto, 0 AS Frete,
	   0 As ValorServico, 0 AS CFOP, '' AS ICMSCST, 0 AS BaseICMS, 0 AS AliquotaICMS, 0 AS ValorICMS,
	   0 AS CodMunOrigem, 0 As CodMunDestino
FROM empresa
WHERE 1 = 2
) tb
/*
SELECT idconhec AS IDLancamento, TbFornecedores.CPF_CNPJ AS IDParticipante, conhec_modelo AS Modelo, '00' As Situacao, (CASE WHEN ISNUMERIC(conhec_serie) = 0 THEN '000' ELSE dbo.LPAD(conhec_serie, 3, '0') END) AS Serie, conhec_numero As Numero,
       ChaveCTe, DataEmissao, emissao_receb AS DataEntradaReceb, '' AS ChaveCTeRef, conhec_valor AS Valor, conhec_desconto As Desconto, (CASE WHEN conhec_cif_fob = 'CIF' THEN '0' ELSE (CASE WHEN conhec_cif_fob = 'FOB' THEN '1' ELSE '9' END) END) AS Frete,
	   conhec_valor As ValorServico, (CASE WHEN conhec_cfop <> '' THEN (CASE WHEN CAST(conhec_cfop AS int) > 4000 THEN CAST(conhec_cfop AS int) - 4000 ELSE CAST(conhec_cfop AS int) END) ELSE '' END) AS CFOP, (CASE WHEN conhec_base_icms > 0 AND conhec_valor_icms > 0 THEN '000' ELSE '041' END) ICMSCST ,conhec_base_icms AS BaseICMS, conhec_aliquota_icms AS AliquotaICMS, conhec_valor_icms AS ValorICMS,
	   TbCidades.CodigoIBGE AS CodMunOrigem, TbFiliais.CodigoIBGE As CodMunDestino
FROM TbConhecimentoFrete
     INNER JOIN TbFornecedores ON TbConhecimentoFrete.IDFornecedor = TbFornecedores.IDFornecedor
	 INNER JOIN TbCidades ON TbFornecedores.CidadeFornecedor = TbCidades.NomeCidade AND TbFornecedores.UFFornecedor = TbCidades.UF
	 INNER JOIN TbFiliais ON TbConhecimentoFrete.IdFilial = TbFiliais.IdFilial
WHERE Fechado <> 0 AND Cancelado = 0 AND conhec_modelo in ('57', '63', '67') AND emissao_receb BETWEEN @dataini AND @datafim AND TbConhecimentoFrete.IdFilial = @filial) tb
*/
GO

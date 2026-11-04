/***************************
PROJETO...: Online.Info
ANALISTA..: Leandro Santos
DESCRICAO.: Procedure de Processamento de Notas SPED
DATA......: 2022-03-15
ALTERACAO.: 
***************************/
-- Idempotente: cria um placeholder se o objeto ainda nao existir e sempre roda o ALTER com a definicao final.
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF OBJECT_ID('dbo.SPEDNotas') IS NULL
    EXEC('CREATE PROCEDURE dbo.SPEDNotas AS SELECT 1')
GO

ALTER PROCEDURE dbo.SPEDNotas(@filial INT, @dataini DATETIME, @datafim DATETIME) AS
DECLARE @msg VARCHAR(200)
DECLARE @somenteinventario INT

--Desabilita a contagem recuperacao/alteracao de dados
SET NOCOUNT ON

--Valor padrão
SET @msg = ''
SET @somenteinventario = 0

--Opa, tem mensagem, entao retorna em formato de erro  
IF (@msg <> '')
 BEGIN
	RAISERROR(@msg, 16, 1)
	RETURN
 END

IF NOT (object_id('vSPEDNotas') is null)
BEGIN
    DROP TABLE vSPEDNotas
END

SELECT IDLancamento, TipoLancamento, tipoOp, SituacaoNF, DataEmissao, DataMovimento, Hora, Emitente, IDParticipante, CNPJCPF, ChaveNFe, NFeCancelada,
       ModeloNF, NumNF, SerieNF, ValorDesconto, ValorFrete, ValorIPI, ValorSeguro, BaseICMS, ValorICMS, BaseICMSST, ValorICMSST, ValorPIS, ValorCOFINS, ValorOutras, ValorNF, ValorTotal, IdFilial, VL
INTO vSPEDNotas
FROM
(
-- Notas de Entrada Emitidas por Fornecedores
SELECT CodigoNota AS IDLancamento, 1 AS TipoLancamento, 0 AS tipoOp, (CASE WHEN (CASE WHEN ModeloNF = '55' OR ModeloNF = '65' THEN SUBSTRING(ChavedeAcesso,23,3) ELSE '0' END) >= '891' THEN '08' ELSE '00' END) AS SituacaoNF ,DataEmissao, DataSaida As DataMovimento, ISNULL(HoraSaida, dbo.hr()) As Hora, '1' AS Emitente, Fornecedor.CPF AS IDParticipante, Fornecedor.CPF AS CNPJCPF, ChavedeAcesso AS ChaveNFe, 0 AS NFeCancelada, 
       ModeloNF, NumeroNota As NumNF, (CASE WHEN ModeloNF = '55' OR ModeloNF = '65' THEN SUBSTRING(ChavedeAcesso,23,3) ELSE '0' END) AS SerieNF, ValorDesconto, ValorFrete, ValorIpi, 0 AS ValorSeguro, 
       (SELECT SUM(CASE WHEN Right(CST, 2) = '60' OR LEN(CST) = 4 THEN 0 ELSE vBC END) FROM EntradaEstoqueItens WHERE EntradaEstoqueItens.CodigoNota = EntradaEstoque.CodigoNota) AS BaseICMS, 
       (SELECT SUM(CASE WHEN Right(CST, 2) = '60' OR LEN(CST) = 4 THEN 0 ELSE vICMS END) FROM EntradaEstoqueItens WHERE EntradaEstoqueItens.CodigoNota = EntradaEstoque.CodigoNota) AS ValorICMS, 
       0 AS BaseICMSST, 0 AS ValorICMSST, ValorPIS AS ValorPIS, ValorCOFINS AS ValorCOFINS, ValorOutrasDespesas AS ValorOutras, (ValorNota + ValorDesconto) As ValorTotal, ValorNota As ValorNF, 1 AS IdFilial, ValorICMSST AS VL
FROM EntradaEstoque 
      INNER JOIN Fornecedor ON EntradaEstoque.CodigoCorrentista = Fornecedor.CODIGO 
WHERE ModeloNF in ('55', '65') AND DataSaida BETWEEN @dataini AND @datafim AND 1 = (CASE WHEN @somenteinventario = 1 THEN 0 ELSE 1 END)

UNION ALL

--NFe de Entrada e Saída com Emissão Própria
SELECT CodigoNota AS IDLancamento, 2 AS TipoLancamento, TipoDocumento AS tipoOp, (CASE WHEN SUBSTRING(ChavedeAcesso,23,3) >= '891' THEN '08' ELSE '00' END) AS SituacaoNF, DataEmissao, ISNULL(DataSaida, DataEmissao) AS DataMovimento, HoraSaida As Hora, '0' AS Emitente, CNPJ_CPF AS IDParticipante, CNPJ_CPF AS CNPJCPF, ChavedeAcesso AS ChaveNFe, Cancelada AS NFeCancelada, 
       '55' AS ModeloNF, NumeroNota As NumNF, SUBSTRING(ChavedeAcesso,23,3) AS SerieNF, ValorDesconto, ValorFrete, ValorIPI, ValorSeguro, BaseICMS, ValorICMS, BaseICMSST, ValorICMSST, 0 AS ValorPIS, 0 AS ValorCOFINS, ValorOutrasDespesas AS ValorOutras, ValorNota As ValorNF, 
       ((ValorNota - ValorDesconto) + (ValorFrete + ValorSeguro + ValorOutrasDespesas + ValorIPI)) AS ValorTotal, 1 AS IDFilial, 0 AS VL 
FROM NotaFiscal
WHERE Enviada <> 0 AND Cancelada = 0 AND Cancelada = 0 AND DataEmissao BETWEEN @dataini AND @datafim AND 1 = (CASE WHEN @somenteinventario = 1 THEN 0 ELSE 1 END)

UNION ALL

--NFCe
SELECT IdNFProd AS IDLancamento, 3 AS TipoLancamento, (CASE WHEN TipoNF = 'S' THEN 1 ELSE 0 END) AS tipoOp, (CASE WHEN SUBSTRING(NFCeChaveAcesso,23,3) >= '891' THEN '08' ELSE '00' END) AS SituacaoNF, DataEmissao, ISNULL(DataSaidaEntrada, DataEmissao) AS DataMovimento, HoraSaida As Hora, '0' AS Emitente, '' AS IDParticipante, '' AS CNPJCPF, NFCeChaveAcesso AS ChaveNFe, NFCeCancelada AS NFeCancelada,
       '65' AS ModeloNF, NumeNota As NumNF, SUBSTRING(NFCeChaveAcesso,23,3) AS SerieNF, DescontoPromocional As ValorDesconto, Valor_Frete As ValorFrete, Valor_Ipi AS ValorIPI, Valor_Seguro AS ValorSeguro,
       BaseCalc_ICMS AS BaseICMS, Valor_ICMS AS ValorICMS, BaseCalc_ICSM_Subst AS BaseICMSST, Valor_ICMS_Subst AS ValorICMSST, 0 AS ValorPIS, 0 AS ValorCOFINS, OutrasDespesasAces AS ValorOutras, Valor_NF_Prod As ValorNF, 
       ((Valor_NF_Prod - DescontoPromocional) + (Valor_Frete + Valor_Seguro + OutrasDespesasAces + Valor_IPI)) AS ValorTotal, 1 AS IdFilial, 0 AS VL
FROM TbNFCe 
WHERE NFCeEnviada <> 0 AND NFCeCancelada = 0 AND Cancelada = 0 AND DataEmissao BETWEEN @dataini AND @datafim AND 1 = (CASE WHEN @somenteinventario = 1 THEN 0 ELSE 1 END)
) tb
GO

/***************************
PROJETO...: Online.Info
ANALISTA..: Leandro Santos
DESCRICAO.: Procedure de Processamento de Notas SPED
DATA......: 2022-03-16
ALTERACAO.: 
***************************/
-- Idempotente: cria um placeholder se o objeto ainda nao existir e sempre roda o ALTER com a definicao final.
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF OBJECT_ID('dbo.SPEDNotasItens') IS NULL
    EXEC('CREATE PROCEDURE dbo.SPEDNotasItens AS SELECT 1')
GO

ALTER PROCEDURE dbo.SPEDNotasItens(@filial INT, @dataini DATETIME, @datafim DATETIME) AS
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

IF NOT (object_id('vSPEDNotasItens') is null)
BEGIN
    DROP TABLE vSPEDNotasItens
END

SELECT IDLancamento, TipoLancamento, Seq, CodigoProduto, NomeProduto, Quantidade, UND, ValorTotal, IndMov, NCM,
ValorOutras, ValorFrete, ValorDesconto, ICMSCST, CFOP, CodNatOperacao, BaseICMS, AliqICMS, ValorICMS,
BaseICMSST, AliqICMSST, ValorICMSST, IndApuracao, IPICST, CodENQ, BaseIPI, AliqIPI, ValorIPI, PISCST, COFINSCST,
QtdeBasePIS, BasePIS, AliqPIS, ValorPIS, QtdeBaseCOFINS, BaseCOFINS, AliqCOFINS, ValorCOFINS, CodContaContabil, VlAbatimentoNT, VlRedBCICMS, TemCredICMS
INTO vSPEDNotasItens
FROM (
-- Notas de Entrada Emitidas por Fornecedores
SELECT EntradaEstoqueItens.CodigoNota AS IDLancamento, 1 AS TipoLancamento, EntradaEstoqueItens.Item AS Seq, EntradaEstoqueItens.CodigoProduto, Produtos.DESCRICAO AS NomeProduto, EntradaEstoqueItens.QuantidadeComercial AS Quantidade,
Produtos.UNID_MEDIDA AS UND, (EntradaEstoqueItens.QuantidadeComercial * EntradaEstoqueItens.ValorUnitarioComercializacao) AS ValorTotal, 0 AS IndMov, EntradaEstoqueItens.NCM, 0 AS ValorOutras, EntradaEstoqueItens.ValorFrete, EntradaEstoqueItens.ValorDesconto,
(CASE WHEN LEN(EntradaEstoqueItens.CST) < 3 THEN '0' + EntradaEstoqueItens.CST ELSE EntradaEstoqueItens.CST END) AS ICMSCST, EntradaEstoqueItens.CFOP, '' AS CodNatOperacao, (CASE WHEN EntradaEstoqueItens.CST = '060' THEN 0 ELSE EntradaEstoqueItens.vBC END) AS BaseICMS, (CASE WHEN EntradaEstoqueItens.CST = '060' THEN 0 ELSE EntradaEstoqueItens.pICMS END) AS AliqICMS, (CASE WHEN EntradaEstoqueItens.CST = '060' THEN 0 ELSE EntradaEstoqueItens.vICMS END) AS ValorICMS, 
0 AS BaseICMSST, 0 AS AliqICMSST, 0 AS ValorICMSST, 0 AS IndApuracao, '03' AS IPICST, '' AS CodENQ, 0 AS BaseIPI, 0 AS AliqIPI, 0 AS ValorIPI, (CASE WHEN EntradaEstoqueItens.PISCST = '' THEN '01' ELSE EntradaEstoqueItens.PISCST END) AS PISCST, (CASE WHEN EntradaEstoqueItens.COFINSCST = '' THEN '01' ELSE EntradaEstoqueItens.COFINSCST END) AS COFINSCST,
0 AS QtdeBasePIS, 0 AS BasePIS, 0 AS AliqPIS, 0 AS ValorPIS, 0 AS QtdeBaseCOFINS, 0 AS BaseCOFINS, 0 AS AliqCOFINS, 0 AS ValorCOFINS, '' AS CodContaContabil, 0 AS VlAbatimentoNT, (CASE WHEN (EntradaEstoqueItens.vBC - (EntradaEstoqueItens.QuantidadeComercial * EntradaEstoqueItens.ValorUnitarioComercializacao)) < 0 THEN ((EntradaEstoqueItens.QuantidadeComercial * EntradaEstoqueItens.ValorUnitarioComercializacao) - EntradaEstoqueItens.vBC) ELSE 0 END ) AS VlRedBCICMS, 0 AS TemCredICMS
FROM EntradaEstoqueItens 
     INNER JOIN Produtos ON EntradaEstoqueItens.CodigoProduto = Produtos.Codigo
     INNER JOIN EntradaEstoque ON EntradaEstoqueItens.CodigoNota = EntradaEstoque.CodigoNota 
WHERE EntradaEstoque.ModeloNF in ('55', '65') AND EntradaEstoque.DataSaida BETWEEN @dataini AND @datafim AND 1 = (CASE WHEN @somenteinventario = 1 THEN 0 ELSE 1 END)

UNION ALL

--Notas Fiscais Eletrônica Emitidas
SELECT NotaFiscalItens.CodigoNota As IDLancamento, 2 AS TipoLancamento, NotaFiscalItens.Item AS Seq, NotaFiscalItens.CodigoProduto, NotaFiscalItens.NomeProduto, NotaFiscalItens.QuantidadeComercial AS Quantidade, 
NotaFiscalItens.UnidadeComercial AS UND, NotaFiscalItens.ValorTotalBruto AS ValorTotal, 0 AS IndMov, NotaFiscalItens.NCM, NotaFiscalItens.ValorOutros AS ValorOutras, NotaFiscalItens.ValorFrete, NotaFiscalItens.ValorDesconto, (CASE WHEN LEN(NotaFiscalItens.CST) < 3 THEN '0' + NotaFiscalItens.CST ELSE NotaFiscalItens.CST END) AS ICMSCST, NotaFiscalItens.CFOP, '' AS CodNatOperacao, 
NotaFiscalItens.vBC AS BaseICMS, NotaFiscalItens.pICMS AS AliqICMS, NotaFiscalItens.vICMS AS ValorICMS, 0 AS BaseICMSST, 0 AS AliqICMSST, 0 AS ValorICMSST, 0 AS IndApuracao, 
NotaFiscalItens.IPICST, '' AS CodENQ, NotaFiscalItens.IPIvBC AS BaseIPI, NotaFiscalItens.IPIpIPI AS Aliq_Ipi, NotaFiscalItens.IPIvIPI AS ValorIPI, NotaFiscalItens.PISCST, NotaFiscalItens.COFINSCST,
0 AS QtdeBasePIS, NotaFiscalItens.PISvBC AS BasePIS, NotaFiscalItens.PISpPIS AS AliqPIS, NotaFiscalItens.PISvPIS AS ValorPIS, 0 AS QtdeBaseCOFINS, NotaFiscalItens.COFINSvBC AS BaseCOFINS, NotaFiscalItens.COFINSpCOFINS AS AliqCOFINS, NotaFiscalItens.COFINSvCOFINS AS ValorCOFINS, '' AS CodContaContabil, 0 AS VlAbatimentoNT, 0 AS VlRedBCICMS, 0 AS TemCredICMS
FROM NotaFiscalItens 
     INNER JOIN NotaFiscal ON NotaFiscalItens.CodigoNota = NotaFiscal.CodigoNota
WHERE NotaFiscal.Enviada <> 0 AND NotaFiscal.Cancelada = 0 AND NotaFiscal.DataEmissao BETWEEN @dataini AND @datafim AND 1 = (CASE WHEN @somenteinventario = 1 THEN 0 ELSE 1 END)

UNION ALL

--Notas Fiscais de Consumidor Eletrônica Emitidas
SELECT TbNFCe_Itens.IdNFProd As IDLancamento, 3 AS TipoLancamento, TbNFCe_Itens.IdNFProd_Item AS Seq, TbNFCe_Itens.CodProduto AS CodigoProduto, TbNFCe_Itens.DescricaoProduto AS NomeProduto, TbNFCe_Itens.QtdeMov AS Quantidade, 
TbNFCe_Itens.UN AS UND, (TbNFCe_Itens.QtdeMov * TbNFCe_Itens.ValorUnit) AS ValorTotal, 0 AS IndMov, CodNCM AS NCM, TbNFCe_Itens.ValorOutras AS ValorOutras, TbNFCe_Itens.Valor_Frete AS ValorFrete, TbNFCe_Itens.Desconto As ValorDesconto, 
(CASE WHEN LEN(TbNFCe_Itens.ICMSCST) < 3 THEN '0' + TbNFCe_Itens.ICMSCST ELSE TbNFCe_Itens.ICMSCST END) AS ICMSCST, TbNFCe_Itens.CFOP, '' AS CodNatOperacao, TbNFCe_Itens.Bc_Icms AS BaseICMS, TbNFCe_Itens.Aliq_Icms AS AliqICMS, TbNFCe_Itens.Vlr_Icms AS ValorICMS,
0 AS BaseICMSST, 0 AS AliqICMSST, 0 AS ValorICMSST, 0 AS IndApuracao, TbNFCe_Itens.IPICST, '' AS CodENQ, (CASE WHEN TbNFCe_Itens.Vlr_IPI > 0 THEN (TbNFCe_Itens.QtdeMov * TbNFCe_Itens.ValorUnit) ELSE 0 END) AS BaseIPI, TbNFCe_Itens.Aliq_Ipi AS AliqIPI, TbNFCe_Itens.Vlr_IPI AS ValorIPI, TbNFCe_Itens.PISCST, TbNFCe_Itens.COFINSCST,
0 AS QtdeBasePIS, 0 AS BasePIS, 0 AS AliqPIS, 0 AS ValorPIS, 0 AS QtdeBaseCOFINS, 0 AS BaseCOFINS, 0 AS AliqCOFINS, 0 AS ValorCOFINS, '' AS CodContaContabil, 0 AS VlAbatimentoNT, 0 AS VlRedBCICMS, 0 AS TemCredICMS
FROM TbNFCe_Itens
     INNER JOIN TbNFCe ON  TbNFCe_Itens.IdNFProd = TbNFCe.IdNFProd
WHERE TbNFCe.NFCeEnviada <> 0 AND TbNFCe.NFCeCancelada = 0 AND TbNFCe.Cancelada = 0 AND TbNFCe.DataEmissao BETWEEN @dataini AND @datafim AND 1 = (CASE WHEN @somenteinventario = 1 THEN 0 ELSE 1 END)
) tb
GO

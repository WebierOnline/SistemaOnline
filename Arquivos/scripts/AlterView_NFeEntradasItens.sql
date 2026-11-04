/***************************
PROJETO...: Online.Info
ANALISTA..: Leandro Santos
DESCRICAO.: View de Nota Fiscal Entradas usada para SPED
DATA......: 2022-03-15
ALTERACAO.: 
***************************/
-- Idempotente: cria um placeholder se o objeto ainda nao existir e sempre roda o ALTER com a definicao final.
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF OBJECT_ID('dbo.NFeEntradasItens') IS NULL
    EXEC('CREATE VIEW dbo.NFeEntradasItens AS SELECT 1 AS Placeholder')
GO

ALTER VIEW dbo.NFeEntradasItens AS
SELECT EntradaEstoqueItens.CodigoNota AS IDLancamento, EntradaEstoqueItens.CodigoProduto AS IDProduto, Produtos.COD_BARRA AS EAN, 
Produtos.NCM, Produtos.Descricao AS NomeProduto, EntradaEstoqueItens.CFOP, Produtos.UNID_MEDIDA AS UND, EntradaEstoqueItens.QuantidadeComercial AS Quantidade, 
EntradaEstoqueItens.ValorUnitarioComercializacao AS ValorUnit, (EntradaEstoqueItens.QuantidadeComercial * EntradaEstoqueItens.ValorUnitarioComercializacao) As ValorTotal, 
EntradaEstoqueItens.ValorFrete, Produtos.ICMSCST AS ICMSCST2, EntradaEstoqueItens.ValorSeguro, EntradaEstoqueItens.ValorDesconto, 
0 AS ValorOutras, EntradaEstoqueItens.CST AS ICMSCST, EntradaEstoqueItens.PISCST, EntradaEstoqueItens.COFINSCST, EntradaEstoqueItens.IPICST, 
EntradaEstoqueItens.vBC AS BaseICMS, EntradaEstoqueItens.pIcms AS AliqICMS, EntradaEstoqueItens.vICMS AS ValorICMS, 
EntradaEstoqueItens.vBCST AS BaseICMSST, EntradaEstoqueItens.vICMSST AS ValorICMSST, EntradaEstoqueItens.IPIpIPI AS Aliq_Ipi, EntradaEstoqueItens.IPIvIPI AS ValorIPI 
FROM EntradaEstoqueItens 
     INNER JOIN Produtos ON EntradaEstoqueItens.CodigoProduto = Produtos.Codigo      
     INNER JOIN EntradaEstoque ON EntradaEstoqueItens.CodigoNota = EntradaEstoque.CodigoNota
GO

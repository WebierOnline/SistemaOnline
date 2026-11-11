/***************************
PROJETO...: Online.Info
ANALISTA..: Leandro Santos
DESCRICAO.: Procedure de Processamento de Notas SPED
DATA......: 2022-03-16
ALTERACAO.: 2026-11 - Notas de entrada: itens sem produto vinculado (uso e consumo lancado sem vincular) entram com o
            codigo 'NE<nota>-<item>' (0200 gerado pelo SpedGerador com a descricao do XML); antes o INNER JOIN com
            produtos tirava o item e a nota saia sem C170/C190. CFOP do item convertido para o de entrada
            (dbo.CFOPEntradaSPED: 5/6/7 -> 1/2/3, venda -> compra para revenda/uso e consumo/imobilizado, com ou sem ST).
            Unidade vazia vira 'UN' (igual ao 0200). CodigoProduto passa a ser texto.
            2026-11 - Produto do item da nota pelo que foi LANCADO no estoque (dbo.SPEDItensEntrada: produtos_entrada pela
            nota ou pela chave de acesso + produtos_entrada_itens.ItemXML); EntradaEstoqueItens fica so com o XML.
            QtdEstoque/UNXML/EANXML para o fator de conversao (0220) quando a nota vem em caixa e o estoque em unidade.
***************************/
-- Idempotente: cria um placeholder se o objeto ainda nao existir e sempre roda o ALTER com a definicao final.
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- CFOP de entrada a partir do CFOP do fornecedor (o XML traz o CFOP de saida dele).
-- @destino: 'R' = revenda (mercadoria), 'U' = uso e consumo, 'I' = ativo imobilizado.
-- ST: CFOP 5401/5402/5403/5405 do fornecedor, CST 10/30/60/70 ou CSOSN 201/202/203/500.
IF OBJECT_ID('dbo.CFOPEntradaSPED') IS NULL
    EXEC('CREATE FUNCTION dbo.CFOPEntradaSPED(@cfop VARCHAR(10)) RETURNS VARCHAR(4) AS BEGIN RETURN @cfop END')
GO
ALTER FUNCTION dbo.CFOPEntradaSPED(@cfop VARCHAR(10), @cst VARCHAR(10), @destino CHAR(1)) RETURNS VARCHAR(4) AS
BEGIN
    DECLARE @c VARCHAR(10), @d CHAR(1), @op CHAR(3), @trib VARCHAR(3), @st BIT, @novo CHAR(3)

    SET @c = LTRIM(RTRIM(ISNULL(@cfop, '')))
    IF LEN(@c) <> 4 OR @c LIKE '%[^0-9]%' RETURN @c

    SET @d = CASE LEFT(@c, 1) WHEN '5' THEN '1' WHEN '6' THEN '2' WHEN '7' THEN '3' ELSE LEFT(@c, 1) END
    SET @op = RIGHT(@c, 3)
    SET @destino = ISNULL(@destino, 'R')

    -- CST com origem (3 digitos): tributacao nos 2 ultimos; CSOSN com origem (4 digitos): os 3 ultimos.
    SET @trib = LTRIM(RTRIM(ISNULL(@cst, '')))
    SET @trib = CASE WHEN LEN(@trib) = 4 THEN RIGHT(@trib, 3) WHEN LEN(@trib) = 3 THEN RIGHT(@trib, 2) ELSE @trib END
    SET @st = CASE WHEN @op IN ('401', '402', '403', '405') OR @trib IN ('10', '30', '60', '70', '201', '202', '203', '500') THEN 1 ELSE 0 END

    SET @novo =
        CASE
            -- vendas do fornecedor = compras
            WHEN @op IN ('101', '102', '103', '104', '105', '106', '107', '108', '109', '110', '111', '113', '114', '115', '116', '117', '118', '119',
                         '120', '122', '123', '401', '402', '403', '405', '929') THEN
                CASE @destino
                    WHEN 'U' THEN CASE WHEN @st = 1 THEN '407' ELSE '556' END
                    WHEN 'I' THEN CASE WHEN @st = 1 THEN '406' ELSE '551' END
                    ELSE CASE WHEN @st = 1 THEN '403' ELSE '102' END
                END
            -- combustiveis
            WHEN @op IN ('651', '652', '653', '654', '655', '656', '667') THEN CASE WHEN @destino = 'R' THEN '652' ELSE '653' END
            -- devolucao de compra do fornecedor = devolucao de venda nossa
            WHEN @op IN ('201', '202') THEN @op
            WHEN @op IN ('410', '411') THEN '411'
            ELSE @op   -- 152, 409, 901, 902, 910, 911, 915, 916, 949... mesmo codigo do lado da entrada
        END

    RETURN @d + @novo
END
GO

-- Item do XML (EntradaEstoqueItens) x item lancado no estoque (produtos_entrada_itens), das notas de entrada do periodo.
-- Entrada lancada: produtos_entrada da mesma nota (CodigoNota) ou da mesma chave de acesso; item pelo ItemXML.
-- Sem lancamento: produto vinculado direto no item do XML (vinculo antigo), senao NULL (uso e consumo).
-- QtdEstoque = quantidade que entrou no estoque, na unidade do produto (fator = QtdEstoque / quantidade do XML).
IF OBJECT_ID('dbo.SPEDItensEntrada') IS NULL
    EXEC('CREATE FUNCTION dbo.SPEDItensEntrada(@dataini DATETIME) RETURNS TABLE AS RETURN SELECT 1 AS x')
GO
ALTER FUNCTION dbo.SPEDItensEntrada(@dataini DATETIME, @datafim DATETIME) RETURNS TABLE AS RETURN
SELECT e.CodigoNota, e.Item,
       COALESCE(l.CodigoProduto, pv.CODIGO) AS IDProduto,
       CAST(CASE WHEN l.CodigoProduto IS NOT NULL THEN l.Qtd ELSE e.QuantidadeComercial END AS DECIMAL(18, 6)) AS QtdEstoque,
       CASE WHEN l.CodigoProduto IS NOT NULL THEN 1 ELSE 0 END AS Lancado
FROM EntradaEstoqueItens e
     INNER JOIN EntradaEstoque ee ON ee.CodigoNota = e.CodigoNota
     OUTER APPLY (SELECT TOP 1 pe.CODIGO FROM produtos_entrada pe
                  WHERE (pe.CodigoNota = ee.CodigoNota AND pe.CodigoNota <> 0)
                     OR (ISNULL(ee.ChavedeAcesso, '') <> '' AND pe.ChavedeAcesso = ee.ChavedeAcesso)
                  ORDER BY CASE WHEN pe.CodigoNota = ee.CodigoNota THEN 0 ELSE 1 END, pe.CODIGO DESC) pe
     OUTER APPLY (SELECT TOP 1 pi.CodigoProduto, SUM(pi.QuantidadeTributavel) AS Qtd FROM produtos_entrada_itens pi
                  INNER JOIN produtos p ON p.CODIGO = pi.CodigoProduto
                  WHERE pi.CODIGO_ENTRADA = pe.CODIGO AND pi.ItemXML = e.Item
                  GROUP BY pi.CodigoProduto ORDER BY SUM(pi.QuantidadeTributavel) DESC) l
     LEFT JOIN produtos pv ON pv.CODIGO = e.CodigoProduto AND e.CodigoProduto <> 0
WHERE ee.ModeloNF IN ('55', '65') AND ee.DataSaida BETWEEN @dataini AND @datafim
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
QtdeBasePIS, BasePIS, AliqPIS, ValorPIS, QtdeBaseCOFINS, BaseCOFINS, AliqCOFINS, ValorCOFINS, CodContaContabil, VlAbatimentoNT, VlRedBCICMS, TemCredICMS, ProdutoAvulso, QtdEstoque, UNXML, EANXML
INTO vSPEDNotasItens
FROM (
-- Notas de Entrada Emitidas por Fornecedores (item sem produto vinculado = uso e consumo, codigo NE<nota>-<item>)
SELECT EntradaEstoqueItens.CodigoNota AS IDLancamento, 1 AS TipoLancamento, EntradaEstoqueItens.Item AS Seq,
CAST(CASE WHEN Produtos.CODIGO IS NULL THEN 'NE' + CAST(EntradaEstoqueItens.CodigoNota AS VARCHAR(20)) + '-' + CAST(EntradaEstoqueItens.Item AS VARCHAR(10)) ELSE CAST(Produtos.CODIGO AS VARCHAR(20)) END AS VARCHAR(60)) AS CodigoProduto,
CAST(CASE WHEN Produtos.CODIGO IS NULL THEN ISNULL(CAST(EntradaEstoqueItens.NomeProduto AS NVARCHAR(250)), N'') ELSE Produtos.DESCRICAO END AS NVARCHAR(250)) AS NomeProduto, EntradaEstoqueItens.QuantidadeComercial AS Quantidade,
CAST(CASE WHEN Produtos.CODIGO IS NULL THEN UPPER(LEFT(ISNULL(NULLIF(LTRIM(RTRIM(EntradaEstoqueItens.UnidadeComercial)), ''), 'UN'), 6)) ELSE ISNULL(NULLIF(LTRIM(RTRIM(Produtos.UNID_MEDIDA)), ''), 'UN') END AS VARCHAR(10)) AS UND,
(EntradaEstoqueItens.QuantidadeComercial * EntradaEstoqueItens.ValorUnitarioComercializacao) AS ValorTotal, 0 AS IndMov, EntradaEstoqueItens.NCM, 0 AS ValorOutras, EntradaEstoqueItens.ValorFrete, EntradaEstoqueItens.ValorDesconto,
(CASE WHEN LEN(EntradaEstoqueItens.CST) < 3 THEN '0' + EntradaEstoqueItens.CST ELSE EntradaEstoqueItens.CST END) AS ICMSCST,
dbo.CFOPEntradaSPED(EntradaEstoqueItens.CFOP, EntradaEstoqueItens.CST,
    CASE WHEN Produtos.CODIGO IS NULL THEN 'U' WHEN ISNULL(Produtos.IMOBILIZADO, 0) <> 0 THEN 'I' WHEN ISNULL(Produtos.USOCONSUMO, 0) <> 0 THEN 'U' ELSE 'R' END) AS CFOP,
'' AS CodNatOperacao, (CASE WHEN EntradaEstoqueItens.CST = '060' THEN 0 ELSE EntradaEstoqueItens.vBC END) AS BaseICMS, (CASE WHEN EntradaEstoqueItens.CST = '060' THEN 0 ELSE EntradaEstoqueItens.pICMS END) AS AliqICMS, (CASE WHEN EntradaEstoqueItens.CST = '060' THEN 0 ELSE EntradaEstoqueItens.vICMS END) AS ValorICMS,
0 AS BaseICMSST, 0 AS AliqICMSST, 0 AS ValorICMSST, 0 AS IndApuracao, '03' AS IPICST, '' AS CodENQ, 0 AS BaseIPI, 0 AS AliqIPI, 0 AS ValorIPI, (CASE WHEN EntradaEstoqueItens.PISCST = '' THEN '01' ELSE EntradaEstoqueItens.PISCST END) AS PISCST, (CASE WHEN EntradaEstoqueItens.COFINSCST = '' THEN '01' ELSE EntradaEstoqueItens.COFINSCST END) AS COFINSCST,
0 AS QtdeBasePIS, 0 AS BasePIS, 0 AS AliqPIS, 0 AS ValorPIS, 0 AS QtdeBaseCOFINS, 0 AS BaseCOFINS, 0 AS AliqCOFINS, 0 AS ValorCOFINS, '' AS CodContaContabil, 0 AS VlAbatimentoNT, (CASE WHEN (EntradaEstoqueItens.vBC - (EntradaEstoqueItens.QuantidadeComercial * EntradaEstoqueItens.ValorUnitarioComercializacao)) < 0 THEN ((EntradaEstoqueItens.QuantidadeComercial * EntradaEstoqueItens.ValorUnitarioComercializacao) - EntradaEstoqueItens.vBC) ELSE 0 END ) AS VlRedBCICMS, 0 AS TemCredICMS,
CASE WHEN Produtos.CODIGO IS NULL THEN 1 ELSE 0 END AS ProdutoAvulso,
v.QtdEstoque, CAST(UPPER(LEFT(ISNULL(NULLIF(LTRIM(RTRIM(EntradaEstoqueItens.UnidadeComercial)), ''), 'UN'), 6)) AS VARCHAR(10)) AS UNXML,
CAST(ISNULL(EntradaEstoqueItens.EAN, '') AS VARCHAR(20)) AS EANXML
FROM dbo.SPEDItensEntrada(@dataini, @datafim) v
     INNER JOIN EntradaEstoqueItens ON EntradaEstoqueItens.CodigoNota = v.CodigoNota AND EntradaEstoqueItens.Item = v.Item
     LEFT JOIN Produtos ON Produtos.CODIGO = v.IDProduto
WHERE 1 = (CASE WHEN @somenteinventario = 1 THEN 0 ELSE 1 END)

UNION ALL

--Notas Fiscais Eletrônica Emitidas
SELECT NotaFiscalItens.CodigoNota As IDLancamento, 2 AS TipoLancamento, NotaFiscalItens.Item AS Seq, CAST(NotaFiscalItens.CodigoProduto AS VARCHAR(60)) AS CodigoProduto, NotaFiscalItens.NomeProduto, NotaFiscalItens.QuantidadeComercial AS Quantidade,
NotaFiscalItens.UnidadeComercial AS UND, NotaFiscalItens.ValorTotalBruto AS ValorTotal, 0 AS IndMov, NotaFiscalItens.NCM, NotaFiscalItens.ValorOutros AS ValorOutras, NotaFiscalItens.ValorFrete, NotaFiscalItens.ValorDesconto, (CASE WHEN LEN(NotaFiscalItens.CST) < 3 THEN '0' + NotaFiscalItens.CST ELSE NotaFiscalItens.CST END) AS ICMSCST, CAST(NotaFiscalItens.CFOP AS VARCHAR(4)) AS CFOP, '' AS CodNatOperacao,
NotaFiscalItens.vBC AS BaseICMS, NotaFiscalItens.pICMS AS AliqICMS, NotaFiscalItens.vICMS AS ValorICMS, 0 AS BaseICMSST, 0 AS AliqICMSST, 0 AS ValorICMSST, 0 AS IndApuracao,
NotaFiscalItens.IPICST, '' AS CodENQ, NotaFiscalItens.IPIvBC AS BaseIPI, NotaFiscalItens.IPIpIPI AS Aliq_Ipi, NotaFiscalItens.IPIvIPI AS ValorIPI, NotaFiscalItens.PISCST, NotaFiscalItens.COFINSCST,
0 AS QtdeBasePIS, NotaFiscalItens.PISvBC AS BasePIS, NotaFiscalItens.PISpPIS AS AliqPIS, NotaFiscalItens.PISvPIS AS ValorPIS, 0 AS QtdeBaseCOFINS, NotaFiscalItens.COFINSvBC AS BaseCOFINS, NotaFiscalItens.COFINSpCOFINS AS AliqCOFINS, NotaFiscalItens.COFINSvCOFINS AS ValorCOFINS, '' AS CodContaContabil, 0 AS VlAbatimentoNT, 0 AS VlRedBCICMS, 0 AS TemCredICMS,
0 AS ProdutoAvulso, CAST(NotaFiscalItens.QuantidadeComercial AS DECIMAL(18, 6)) AS QtdEstoque, CAST(NotaFiscalItens.UnidadeComercial AS VARCHAR(10)) AS UNXML, CAST('' AS VARCHAR(20)) AS EANXML
FROM NotaFiscalItens
     INNER JOIN NotaFiscal ON NotaFiscalItens.CodigoNota = NotaFiscal.CodigoNota
WHERE NotaFiscal.Enviada <> 0 AND NotaFiscal.Cancelada = 0 AND NotaFiscal.DataEmissao BETWEEN @dataini AND @datafim AND 1 = (CASE WHEN @somenteinventario = 1 THEN 0 ELSE 1 END)

UNION ALL

--Notas Fiscais de Consumidor Eletrônica Emitidas
SELECT TbNFCe_Itens.IdNFProd As IDLancamento, 3 AS TipoLancamento, TbNFCe_Itens.IdNFProd_Item AS Seq, CAST(TbNFCe_Itens.CodProduto AS VARCHAR(60)) AS CodigoProduto, TbNFCe_Itens.DescricaoProduto AS NomeProduto, TbNFCe_Itens.QtdeMov AS Quantidade,
TbNFCe_Itens.UN AS UND, (TbNFCe_Itens.QtdeMov * TbNFCe_Itens.ValorUnit) AS ValorTotal, 0 AS IndMov, CodNCM AS NCM, TbNFCe_Itens.ValorOutras AS ValorOutras, TbNFCe_Itens.Valor_Frete AS ValorFrete, TbNFCe_Itens.Desconto As ValorDesconto,
(CASE WHEN LEN(TbNFCe_Itens.ICMSCST) < 3 THEN '0' + TbNFCe_Itens.ICMSCST ELSE TbNFCe_Itens.ICMSCST END) AS ICMSCST, CAST(TbNFCe_Itens.CFOP AS VARCHAR(4)) AS CFOP, '' AS CodNatOperacao, TbNFCe_Itens.Bc_Icms AS BaseICMS, TbNFCe_Itens.Aliq_Icms AS AliqICMS, TbNFCe_Itens.Vlr_Icms AS ValorICMS,
0 AS BaseICMSST, 0 AS AliqICMSST, 0 AS ValorICMSST, 0 AS IndApuracao, TbNFCe_Itens.IPICST, '' AS CodENQ, (CASE WHEN TbNFCe_Itens.Vlr_IPI > 0 THEN (TbNFCe_Itens.QtdeMov * TbNFCe_Itens.ValorUnit) ELSE 0 END) AS BaseIPI, TbNFCe_Itens.Aliq_Ipi AS AliqIPI, TbNFCe_Itens.Vlr_IPI AS ValorIPI, TbNFCe_Itens.PISCST, TbNFCe_Itens.COFINSCST,
0 AS QtdeBasePIS, 0 AS BasePIS, 0 AS AliqPIS, 0 AS ValorPIS, 0 AS QtdeBaseCOFINS, 0 AS BaseCOFINS, 0 AS AliqCOFINS, 0 AS ValorCOFINS, '' AS CodContaContabil, 0 AS VlAbatimentoNT, 0 AS VlRedBCICMS, 0 AS TemCredICMS,
0 AS ProdutoAvulso, CAST(TbNFCe_Itens.QtdeMov AS DECIMAL(18, 6)) AS QtdEstoque, CAST(TbNFCe_Itens.UN AS VARCHAR(10)) AS UNXML, CAST('' AS VARCHAR(20)) AS EANXML
FROM TbNFCe_Itens
     INNER JOIN TbNFCe ON  TbNFCe_Itens.IdNFProd = TbNFCe.IdNFProd
WHERE TbNFCe.NFCeEnviada <> 0 AND TbNFCe.NFCeCancelada = 0 AND TbNFCe.Cancelada = 0 AND TbNFCe.DataEmissao BETWEEN @dataini AND @datafim AND 1 = (CASE WHEN @somenteinventario = 1 THEN 0 ELSE 1 END)
) tb
GO

/***************************
PROJETO...: Online.Info
DESCRICAO.: Movimento do ESTOQUE FISCAL por produto (base do calculo do saldo fiscal em uma data).
            Quantidade com sinal: + entrada fiscal, - saida fiscal.
            Quantidade = QuantidadeComercial; se zero (as telas preenchem uma ou outra), QuantidadeTributavel.
            ENTRADA  = produtos_entrada(_itens): com chave de acesso e nao cancelada (entrada manual sem nota fica fora)
            NFE      = NotaFiscal(Itens): com protocolo de autorizacao e nao cancelada (inclui devolucao de compra, que e saida)
            NFCE     = TbNFCe(_Itens): com protocolo, sem cancelamento
            AJUSTE   = TbAjusteFiscal: ajustes manuais do estoque fiscal (EstoqueFiscalAjustar)
DATA......: 2026-11
***************************/
-- Idempotente: cria um placeholder se o objeto ainda nao existir e sempre roda o ALTER com a definicao final.
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF OBJECT_ID('dbo.vMovimentoFiscal') IS NULL
    EXEC('CREATE VIEW dbo.vMovimentoFiscal AS SELECT 1 AS Placeholder')
GO

ALTER VIEW dbo.vMovimentoFiscal AS
SELECT 'ENTRADA' AS Origem,
       i.CodigoProduto AS IdProduto,
       CAST(ISNULL(e.DATA_ENTRADA, e.DataEmissao) AS DATE) AS Data,
       CONVERT(DECIMAL(18,3), COALESCE(NULLIF(i.QuantidadeComercial, 0), NULLIF(i.QuantidadeTributavel, 0), 0)) AS Quantidade,
       e.CODIGO AS IdDocumento,
       CONVERT(VARCHAR(20), e.NumeroNota) AS Numero
FROM dbo.produtos_entrada e
INNER JOIN dbo.produtos_entrada_itens i ON i.CODIGO_ENTRADA = e.CODIGO
WHERE ISNULL(e.Cancelada, 0) = 0
  AND ISNULL(e.ChavedeAcesso, '') <> ''

UNION ALL

SELECT 'NFE',
       i.CodigoProduto,
       CAST(n.DataEmissao AS DATE),
       CONVERT(DECIMAL(18,3), -COALESCE(NULLIF(i.QuantidadeComercial, 0), NULLIF(i.QuantidadeTributavel, 0), 0)),
       n.CodigoNota,
       CONVERT(VARCHAR(20), n.NumeroNota)
FROM dbo.NotaFiscal n
INNER JOIN dbo.NotaFiscalItens i ON i.CodigoNota = n.CodigoNota
WHERE ISNULL(n.Cancelada, 0) = 0
  AND ISNULL(n.NumeroProtocolo, 0) > 0

UNION ALL

SELECT 'NFCE',
       i.IDProduto,
       CAST(n.DataEmissao AS DATE),
       CONVERT(DECIMAL(18,3), -ISNULL(i.QtdeMov, 0)),
       n.IdNFProd,
       CONVERT(VARCHAR(20), n.IdNFProd)
FROM dbo.TbNFCe n
INNER JOIN dbo.TbNFCe_Itens i ON i.IdNFProd = n.IdNFProd
WHERE ISNULL(n.Cancelada, 0) = 0
  AND ISNULL(n.NFCeCancelada, 0) = 0
  AND ISNULL(n.NFCeProtocolo, 0) > 0

UNION ALL

SELECT 'AJUSTE',
       a.IdProduto,
       CAST(a.Data AS DATE),
       a.Quantidade,
       a.IdAjuste,
       CONVERT(VARCHAR(20), a.IdAjuste)
FROM dbo.TbAjusteFiscal a
GO

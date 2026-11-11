/***************************
PROJETO...: Online.Info
ANALISTA..: Leandro
DESCRICAO.: Procedure de Processamento de Produtos SPED
DATA......: 2022-03-15
ALTERACAO.: 2026-11 - filtro unico de uso/consumo e imobilizado (@usoconsumo/@imobilizado = 1 incluem esses produtos;
            o produto avulso (codigo 1) nao entra; antes, com os dois marcados, nenhum ramo rodava e a tabela vSPEDProdCod nao era criada).
            2026-11 - TipoItem do 0200 pelo cadastro: 08 = ativo imobilizado, 07 = uso e consumo, 00 = mercadoria.
            2026-11 - produtos das notas de entrada pelo que foi lancado no estoque (dbo.SPEDItensEntrada, script 156),
            nao mais pelo produto gravado no item do XML; produto sem unidade entra (UN, igual ao 0200).
***************************/
-- Idempotente: cria um placeholder se o objeto ainda nao existir e sempre roda o ALTER com a definicao final.
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF OBJECT_ID('dbo.SPEDProdutos') IS NULL
    EXEC('CREATE PROCEDURE dbo.SPEDProdutos AS SELECT 1')
GO

ALTER PROCEDURE dbo.SPEDProdutos(@filial INT, @dataini DATETIME, @datafim DATETIME, @inventario INT, @usoconsumo BIT, @imobilizado BIT, @tipoescrituracao INT, @somenteinventario INT, @blococ170saidas INT) AS

--Desabilita a contagem recuperacao/alteracao de dados
SET NOCOUNT ON

SET @usoconsumo = ISNULL(@usoconsumo, 0)
SET @imobilizado = ISNULL(@imobilizado, 0)
SET @inventario = ISNULL(@inventario, 0)

IF NOT (object_id('vSPEDProdCod') is null)
BEGIN
    DROP TABLE vSPEDProdCod
END

--Produtos das notas de entrada do periodo (EFD completa) + dos itens do inventario (Bloco H)
SELECT IDProduto
INTO vSPEDProdCod
FROM
(
SELECT IDProduto
FROM dbo.SPEDItensEntrada(@dataini, @datafim)
WHERE @somenteinventario = 0 AND IDProduto IS NOT NULL
UNION ALL
SELECT IDProduto
FROM InventarioGerado
WHERE @inventario <> 0 AND MetaCalculado > 0 AND IdInventario = @inventario AND IDProduto <> 1
      AND (@imobilizado <> 0 OR Imobilizado = 0) AND (@usoconsumo <> 0 OR UsoConsumo = 0)
) tb
GROUP BY IDProduto

IF NOT (object_id('vSPEDProdutos') is null)
BEGIN
    DROP TABLE vSPEDProdutos
END

SELECT CODIGO AS CodigoProduto, DESCRICAO AS NomeProduto, (CASE WHEN LEFT(COD_BARRA, 3) = '789' THEN COD_BARRA ELSE '' END) AS EAN, ISNULL((CASE WHEN UNID_MEDIDA = '' THEN 'UN' ELSE UNID_MEDIDA END), 'UN') AS UN, NCM,
CAST(CASE WHEN ISNULL(IMOBILIZADO, 0) <> 0 THEN '08' WHEN ISNULL(USOCONSUMO, 0) <> 0 THEN '07' ELSE '00' END AS VARCHAR(2)) AS TipoItem, ICMSAliq AS ICMS, CEST
INTO vSPEDProdutos
FROM Produtos
WHERE CODIGO IN (SELECT IDProduto FROM vSPEDProdCod)
GO

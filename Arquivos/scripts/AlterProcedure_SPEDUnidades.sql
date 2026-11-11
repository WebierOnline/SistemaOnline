/***************************
PROJETO...: Online.Info
ANALISTA..: Leandro
DESCRICAO.: Procedure de Processamento de Unidades SPED
DATA......: 2022-03-15
ALTERACAO.: 2026-11 - filtro unico de uso/consumo e imobilizado (@usoconsumo/@imobilizado = 1 incluem esses produtos;
            o produto avulso (codigo 1) nao entra; antes, com os dois marcados, nenhum ramo rodava e a tabela vSPEDUnidades nao era criada).
***************************/
-- Idempotente: cria um placeholder se o objeto ainda nao existir e sempre roda o ALTER com a definicao final.
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF OBJECT_ID('dbo.SPEDUnidades') IS NULL
    EXEC('CREATE PROCEDURE dbo.SPEDUnidades AS SELECT 1')
GO

ALTER PROCEDURE dbo.SPEDUnidades(@filial INT, @dataini DATETIME, @datafim DATETIME, @inventario INT, @usoconsumo BIT, @imobilizado BIT, @tipoescrituracao INT, @somenteinventario INT, @blococ170saidas INT) AS

--Desabilita a contagem recuperacao/alteracao de dados
SET NOCOUNT ON

SET @usoconsumo = ISNULL(@usoconsumo, 0)
SET @imobilizado = ISNULL(@imobilizado, 0)
SET @inventario = ISNULL(@inventario, 0)

IF NOT (object_id('vSPEDUnidades') is null)
BEGIN
    DROP TABLE vSPEDUnidades
END

--Unidades das notas de entrada do periodo (EFD completa) + dos itens do inventario (Bloco H)
SELECT UND
INTO vSPEDUnidades
FROM
(
SELECT UND
FROM NFeEntradasItens
WHERE @somenteinventario = 0 AND UND <> '' AND IDLancamento in (SELECT CodigoNota FROM EntradaEstoque WHERE EntradaEstoque.ModeloNF = '55' AND EntradaEstoque.DataSaida BETWEEN @dataini AND @datafim)
UNION ALL
SELECT UN AS UND
FROM InventarioGerado
WHERE @inventario <> 0 AND MetaCalculado > 0 AND IdInventario = @inventario AND IDProduto <> 1
      AND (@imobilizado <> 0 OR Imobilizado = 0) AND (@usoconsumo <> 0 OR UsoConsumo = 0)
) tb
GROUP BY UND
GO

/***************************
PROJETO...: Online.Info
ANALISTA..: Leandro Santos
DESCRICAO.: View do itens do Inventário Gerado
DATA......: 2022-03-16
ALTERACAO.: 2026-11 - Imobilizado e UsoConsumo vem do cadastro do produto (Produtos.IMOBILIZADO/USOCONSUMO),
            antes fixos em 0: o SPED separa esses produtos do inventario de mercadorias.
***************************/
-- Idempotente: cria um placeholder se o objeto ainda nao existir e sempre roda o ALTER com a definicao final.
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF OBJECT_ID('dbo.InventarioGerado') IS NULL
    EXEC('CREATE VIEW dbo.InventarioGerado AS SELECT 1 AS Placeholder')
GO

ALTER VIEW dbo.InventarioGerado AS
SELECT ii.IdInventario, Seq, ii.IDProduto, ii.NomeProduto, ISNULL((CASE WHEN UNID_MEDIDA = '' THEN 'UN' ELSE UNID_MEDIDA END), 'UN') AS UN, (CASE WHEN LEFT(COD_BARRA, 3) = '789' THEN COD_BARRA ELSE '' END) AS EAN, ISNULL(p.NCM, '') AS NCM, p.ICMSAliq AS Aliquota,
       SaldoCalculado, VlrUnitInvent, TotalInvent, MetaCalculado, SaldoLancado, (VlrUnitInvent * SaldoLancado) AS TotalFisico, i.DataInventario, i.IdFilial,
       (CASE WHEN ISNULL(p.IMOBILIZADO, 0) <> 0 THEN 1 ELSE 0 END) AS Imobilizado, (CASE WHEN ISNULL(p.USOCONSUMO, 0) <> 0 THEN 1 ELSE 0 END) AS UsoConsumo,
	   0 AS MateriaPrima, 0 AS ProdutoIndustrializado, i.AnoInventario, 0 AS H020_Aliquota, p.ICMSCST, 0 AS H020_ValorICMS
FROM TbInventariosItens ii
     INNER JOIN TbInventarios i ON ii.IdInventario = i.IdInventario
     LEFT OUTER JOIN Produtos p ON ii.IDProduto = p.Codigo
GO

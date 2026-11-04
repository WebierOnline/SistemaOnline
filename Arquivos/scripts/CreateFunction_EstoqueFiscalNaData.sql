/***************************
PROJETO...: Online.Info
DESCRICAO.: Saldo do ESTOQUE FISCAL de cada produto em uma data (fim do dia @data).
            Saldo = saldo-base + movimento fiscal (vMovimentoFiscal) depois da base, ate @data.
            Saldo-base = SaldoLancado do ultimo inventario FECHADO (Fechado = 1) com DataInventario <= @data.
            Sem inventario fechado ate a data: base 0 e o movimento conta desde o inicio.
            Uso: SELECT * FROM dbo.EstoqueFiscalNaData('20251231')
DATA......: 2026-11
***************************/
-- Idempotente: recria a function se ja existir.
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF OBJECT_ID('dbo.EstoqueFiscalNaData') IS NOT NULL DROP FUNCTION dbo.EstoqueFiscalNaData
GO

CREATE FUNCTION dbo.EstoqueFiscalNaData(@data DATE)
RETURNS TABLE
AS
RETURN
(
    SELECT x.IdProduto,
           CONVERT(DECIMAL(18,3), SUM(x.Quantidade)) AS Saldo,
           MAX(x.Base) AS SaldoBase,
           MAX(x.DataInvBase) AS DataInvBase
    FROM (
        -- saldo-base do ultimo inventario fechado ate a data
        SELECT i.IDProduto AS IdProduto, CONVERT(DECIMAL(18,3), ISNULL(i.SaldoLancado, 0)) AS Quantidade,
               CONVERT(DECIMAL(18,3), ISNULL(i.SaldoLancado, 0)) AS Base, CAST(b.DataInventario AS DATE) AS DataInvBase
        FROM dbo.TbInventariosItens i
        INNER JOIN dbo.TbInventarios b ON b.IdInventario = i.IdInventario
        WHERE b.IdInventario = (SELECT TOP 1 f.IdInventario FROM dbo.TbInventarios f
                                WHERE f.Fechado = 1 AND CAST(f.DataInventario AS DATE) <= @data
                                ORDER BY f.DataInventario DESC, f.IdInventario DESC)

        UNION ALL

        -- movimento fiscal depois da base (ou desde o inicio, sem base) ate a data
        SELECT m.IdProduto, m.Quantidade, CONVERT(DECIMAL(18,3), 0), NULL
        FROM dbo.vMovimentoFiscal m
        WHERE m.Data <= @data
          AND m.Data > ISNULL((SELECT TOP 1 CAST(f.DataInventario AS DATE) FROM dbo.TbInventarios f
                               WHERE f.Fechado = 1 AND CAST(f.DataInventario AS DATE) <= @data
                               ORDER BY f.DataInventario DESC, f.IdInventario DESC), '19000101')
    ) x
    WHERE x.IdProduto IS NOT NULL
    GROUP BY x.IdProduto
)
GO

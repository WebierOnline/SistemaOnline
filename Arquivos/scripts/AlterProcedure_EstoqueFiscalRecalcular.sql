/***************************
PROJETO...: Online.Info
DESCRICAO.: Recalcula Produtos.ESTOQUE_FISCAL (saldo atual) = EstoqueFiscalNaData(hoje).
            So roda se existir ao menos um inventario FECHADO (a abertura); sem base o calculo daria saldos negativos.
DATA......: 2026-11
***************************/
-- Idempotente: cria um placeholder se o objeto ainda nao existir e sempre roda o ALTER com a definicao final.
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF OBJECT_ID('dbo.EstoqueFiscalRecalcular') IS NULL
    EXEC('CREATE PROCEDURE dbo.EstoqueFiscalRecalcular AS SELECT 1')
GO

ALTER PROCEDURE dbo.EstoqueFiscalRecalcular AS
SET NOCOUNT ON

IF NOT EXISTS (SELECT 1 FROM TbInventarios WHERE Fechado = 1 AND DataInventario IS NOT NULL AND CAST(DataInventario AS DATE) <= CAST(GETDATE() AS DATE))
 BEGIN
   RAISERROR('Não há Inventário de Abertura (fechado) para servir de base ao estoque fiscal.', 16, 1)
   RETURN
 END

UPDATE p SET ESTOQUE_FISCAL = ISNULL(s.Saldo, 0)
FROM Produtos p
LEFT JOIN dbo.EstoqueFiscalNaData(CAST(GETDATE() AS DATE)) s ON s.IdProduto = p.CODIGO
WHERE ISNULL(p.ESTOQUE_FISCAL, -999999) <> ISNULL(s.Saldo, 0)
GO

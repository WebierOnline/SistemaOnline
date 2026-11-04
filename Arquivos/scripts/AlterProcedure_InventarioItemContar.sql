/***************************
PROJETO...: Online.Info
DESCRICAO.: Contagem fisica: grava a quantidade contada de um item do inventario.
            SaldoLancado e MetaCalculado recebem a contagem (a Meta e o que o SPED le); SaldoCalculado continua
            sendo o saldo do sistema, para mostrar a diferenca. Inventario FECHADO nao aceita contagem.
DATA......: 2026-11
***************************/
-- Idempotente: cria um placeholder se o objeto ainda nao existir e sempre roda o ALTER com a definicao final.
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF OBJECT_ID('dbo.InventarioItemContar') IS NULL
    EXEC('CREATE PROCEDURE dbo.InventarioItemContar AS SELECT 1')
GO

ALTER PROCEDURE dbo.InventarioItemContar(@id INT, @produto INT, @qtd DECIMAL(18,3)) AS
SET NOCOUNT ON

IF NOT EXISTS (SELECT 1 FROM TbInventarios WHERE IdInventario = @id)
 BEGIN
   RAISERROR('Inventário não encontrado!', 16, 1)
   RETURN
 END

IF EXISTS (SELECT 1 FROM TbInventarios WHERE IdInventario = @id AND Fechado = 1)
 BEGIN
   RAISERROR('Inventário FECHADO: reabra para alterar a contagem!', 16, 1)
   RETURN
 END

IF (@qtd IS NULL OR @qtd < 0)
 BEGIN
   RAISERROR('Quantidade contada inválida!', 16, 1)
   RETURN
 END

UPDATE TbInventariosItens
SET SaldoLancado = @qtd,
    MetaCalculado = @qtd,
    TotalInvent = CONVERT(DECIMAL(12,2), @qtd * VlrUnitInvent)
WHERE IdInventario = @id AND IDProduto = @produto

IF (@@ROWCOUNT = 0)
   RAISERROR('Produto não encontrado no inventário!', 16, 1)
GO

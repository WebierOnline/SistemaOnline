/***************************
PROJETO...: Online.Info
DESCRICAO.: Fecha (@fechar = 1) ou reabre (@fechar = 0) um inventario.
            Inventario fechado vira base do estoque fiscal (EstoqueFiscalNaData) e nao aceita mais contagem.
            Ao fechar/reabrir o ESTOQUE_FISCAL dos produtos e recalculado.
DATA......: 2026-11
***************************/
-- Idempotente: cria um placeholder se o objeto ainda nao existir e sempre roda o ALTER com a definicao final.
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF OBJECT_ID('dbo.InventarioFechar') IS NULL
    EXEC('CREATE PROCEDURE dbo.InventarioFechar AS SELECT 1')
GO

ALTER PROCEDURE dbo.InventarioFechar(@id INT, @fechar BIT = 1) AS
DECLARE @trans INT, @msg VARCHAR(200)
SET NOCOUNT ON

IF NOT EXISTS (SELECT 1 FROM TbInventarios WHERE IdInventario = @id)
 BEGIN
   RAISERROR('Inventário não encontrado!', 16, 1)
   RETURN
 END

IF (@fechar = 1 AND NOT EXISTS (SELECT 1 FROM TbInventariosItens WHERE IdInventario = @id))
 BEGIN
   RAISERROR('Inventário sem itens não pode ser fechado!', 16, 1)
   RETURN
 END

SET @trans = @@TRANCOUNT
BEGIN TRY
   IF (@trans = 0) BEGIN TRANSACTION

   UPDATE TbInventarios SET Fechado = @fechar WHERE IdInventario = @id

   --o saldo fiscal passa a partir da nova base (se ainda houver alguma)
   IF EXISTS (SELECT 1 FROM TbInventarios WHERE Fechado = 1 AND DataInventario IS NOT NULL AND CAST(DataInventario AS DATE) <= CAST(GETDATE() AS DATE))
      EXEC dbo.EstoqueFiscalRecalcular

   IF (@trans = 0) COMMIT TRANSACTION
END TRY
BEGIN CATCH
   SET @msg = ERROR_MESSAGE()
   IF (@trans = 0 AND @@TRANCOUNT > 0) ROLLBACK TRANSACTION
   RAISERROR(@msg, 16, 1)
END CATCH
GO

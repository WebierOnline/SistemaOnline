/***************************
PROJETO...: Online.Info
DESCRICAO.: Ajuste manual do ESTOQUE FISCAL de um produto para o valor @novo.
            Com inventario fechado (Abertura) como base: a diferenca entre @novo e o saldo calculado de hoje vira
            movimento (TbAjusteFiscal) para o ajuste sobreviver ao EstoqueFiscalRecalcular; se a base e de hoje,
            o ajuste entra no proprio item da base (movimento do dia da base nao conta depois dela).
            Sem base: so grava Produtos.ESTOQUE_FISCAL (comportamento antigo).
DATA......: 2026-11
***************************/
-- Idempotente: cria um placeholder se o objeto ainda nao existir e sempre roda o ALTER com a definicao final.
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF OBJECT_ID('dbo.EstoqueFiscalAjustar') IS NULL
    EXEC('CREATE PROCEDURE dbo.EstoqueFiscalAjustar AS SELECT 1')
GO

ALTER PROCEDURE dbo.EstoqueFiscalAjustar(@cod INT, @novo DECIMAL(18,3), @usuario INT = 0) AS
DECLARE @hoje DATE, @idBase INT, @dtBase DATE, @atual DECIMAL(18,3), @delta DECIMAL(18,3), @trans INT, @msg VARCHAR(200)

SET NOCOUNT ON
SET @hoje = CAST(GETDATE() AS DATE)

SELECT TOP 1 @idBase = IdInventario, @dtBase = CAST(DataInventario AS DATE)
FROM TbInventarios
WHERE Fechado = 1 AND DataInventario IS NOT NULL AND CAST(DataInventario AS DATE) <= @hoje
ORDER BY DataInventario DESC, IdInventario DESC

SET @trans = @@TRANCOUNT
BEGIN TRY
   IF (@trans = 0) BEGIN TRANSACTION

   IF (@idBase IS NOT NULL)
    BEGIN
      SELECT @atual = ISNULL((SELECT Saldo FROM dbo.EstoqueFiscalNaData(@hoje) WHERE IdProduto = @cod), 0)
      SET @delta = @novo - @atual

      IF (@delta <> 0)
       BEGIN
         IF (@dtBase = @hoje)
          BEGIN
            IF EXISTS (SELECT 1 FROM TbInventariosItens WHERE IdInventario = @idBase AND IDProduto = @cod)
               UPDATE TbInventariosItens
               SET SaldoLancado = SaldoLancado + @delta,
                   TotalInvent = CONVERT(DECIMAL(12,2), (SaldoLancado + @delta) * ISNULL(VlrUnitInvent, 0))
               WHERE IdInventario = @idBase AND IDProduto = @cod
            ELSE
               INSERT INTO TbInventariosItens(IdInventario, Seq, IDProduto, NomeProduto, SaldoCalculado, SaldoLancado, MetaCalculado, VlrUnitInvent, TotalInvent)
               SELECT @idBase, ISNULL((SELECT MAX(Seq) FROM TbInventariosItens WHERE IdInventario = @idBase), 0) + 1, p.CODIGO, ISNULL(p.DESCRICAO, ''),
                      0, @delta, 0, c.Custo, CONVERT(DECIMAL(12,2), @delta * c.Custo)
               FROM Produtos p
               CROSS APPLY (SELECT ISNULL((SELECT TOP 1 pr.CUSTO FROM Produtos_Precos pr WHERE pr.COD_PRODUTO = p.CODIGO ORDER BY pr.CODIGO DESC), 0) AS Custo) c
               WHERE p.CODIGO = @cod
          END
         ELSE
            INSERT INTO TbAjusteFiscal(Data, IdProduto, Quantidade, IdUsuario, Obs)
            VALUES(GETDATE(), @cod, @delta, @usuario, 'Ajuste manual do estoque fiscal')
       END
    END

   UPDATE Produtos SET ESTOQUE_FISCAL = @novo WHERE CODIGO = @cod

   IF (@trans = 0) COMMIT TRANSACTION
END TRY
BEGIN CATCH
   SET @msg = ERROR_MESSAGE()
   IF (@trans = 0 AND @@TRANCOUNT > 0) ROLLBACK TRANSACTION
   RAISERROR(@msg, 16, 1)
END CATCH
GO

/***************************
PROJETO...: Online.Info
DESCRICAO.: Inventario de ABERTURA do estoque fiscal: gera, ja FECHADO, um inventario (tipo 'Abertura') com o
            ESTOQUE_FISCAL atual de cada produto como saldo-base. A function EstoqueFiscalNaData usa o ultimo
            inventario fechado como base e soma o movimento fiscal (vMovimentoFiscal) depois dele.
            Data de referencia = hoje. Entram so os produtos ativos com saldo positivo.
DATA......: 2026-11
***************************/
-- Idempotente: cria um placeholder se o objeto ainda nao existir e sempre roda o ALTER com a definicao final.
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF OBJECT_ID('dbo.InventarioAbertura') IS NULL
    EXEC('CREATE PROCEDURE dbo.InventarioAbertura AS SELECT 1')
GO

ALTER PROCEDURE dbo.InventarioAbertura(@fu INT, @filial INT = 1) AS
DECLARE @idInv INT, @dt DATETIME, @trans INT, @msg VARCHAR(200)

SET NOCOUNT ON
SET @filial = ISNULL(@filial, 1)
SET @dt = CAST(CAST(GETDATE() AS DATE) AS DATETIME)

IF EXISTS (SELECT 1 FROM TbInventarios WHERE PeriodoInventarioTipo = 'Abertura' AND IdFilial = @filial AND DataInventario = @dt)
 BEGIN
   RAISERROR('Já existe um Inventário de Abertura gerado hoje!', 16, 1)
   RETURN
 END

SET @trans = @@TRANCOUNT
BEGIN TRY
   IF (@trans = 0) BEGIN TRANSACTION

   SELECT @idInv = ISNULL(MAX(IdInventario), 0) + 1 FROM TbInventarios WITH (UPDLOCK, HOLDLOCK)

   INSERT INTO TbInventarios(IdInventario, Data, IDFuncionario, AnoInventario, PeriodoInventarioTipo, PeriodoInventario, DataInventario, Fechado, Nada, IdFilial)
   VALUES(@idInv, GETDATE(), @fu, YEAR(@dt), 'Abertura', '', @dt, 1, '', @filial)

   INSERT INTO TbInventariosItens
              (IdInventario, Seq, IDProduto, NomeProduto, SaldoCalculado, SaldoLancado, MetaCalculado, VlrUnitInvent, TotalInvent)
   SELECT @idInv, ROW_NUMBER() OVER(ORDER BY x.DESCRICAO, x.CODIGO), x.CODIGO, x.DESCRICAO, x.Saldo, x.Saldo, x.Saldo,
          x.Custo, CONVERT(DECIMAL(12,2), x.Saldo * x.Custo)
   FROM (SELECT p.CODIGO, ISNULL(p.DESCRICAO, '') AS DESCRICAO, ISNULL(p.ESTOQUE_FISCAL, 0) AS Saldo,
                ISNULL((SELECT TOP 1 pr.CUSTO FROM Produtos_Precos pr WHERE pr.COD_PRODUTO = p.CODIGO ORDER BY pr.CODIGO DESC), 0) AS Custo
         FROM Produtos p
         WHERE (p.ATIVO = 1) AND (ISNULL(p.ESTOQUE_FISCAL, 0) > 0)) x
   ORDER BY x.DESCRICAO, x.CODIGO

   IF (@trans = 0) COMMIT TRANSACTION
END TRY
BEGIN CATCH
   SET @msg = ERROR_MESSAGE()
   IF (@trans = 0 AND @@TRANCOUNT > 0) ROLLBACK TRANSACTION
   RAISERROR(@msg, 16, 1)
END CATCH
GO

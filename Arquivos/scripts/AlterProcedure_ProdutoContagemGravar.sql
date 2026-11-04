/***************************
PROJETO...: Online.Info
DESCRICAO.: Contagem fisica de produtos (tela Produtos_Contagem). O usuario conta na prateleira e informa a quantidade:
            a procedure le o estoque atual, grava o ajuste em Produtos_Quant (FORMA = 'CONTAGEM', ADICAO ou REMOCAO,
            QUANT = diferenca, ESTOQUE = quantidade depois do ajuste), altera Produtos.QUANT_ESTOQUE e marca o produto
            como contado (Produtos.CONTADO / DATA_CONTAGEM). Sem diferenca nao grava ajuste, so marca como contado.
            Tudo na mesma transacao; a quantidade anterior e lida com o produto travado (venda no PDV durante a contagem
            nao se perde). Devolve uma linha: QuantAnterior, QuantNova.
DATA......: 2026-11
***************************/
-- Idempotente: acrescenta as colunas so se faltarem e sempre roda o ALTER com a definicao final.
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

IF COL_LENGTH('dbo.Produtos', 'CONTADO') IS NULL
    ALTER TABLE dbo.Produtos ADD CONTADO BIT NULL
IF COL_LENGTH('dbo.Produtos', 'DATA_CONTAGEM') IS NULL
    ALTER TABLE dbo.Produtos ADD DATA_CONTAGEM DATETIME NULL
GO

IF OBJECT_ID('dbo.ProdutoContagemGravar') IS NULL
    EXEC('CREATE PROCEDURE dbo.ProdutoContagemGravar AS SELECT 1')
GO

ALTER PROCEDURE dbo.ProdutoContagemGravar(@produto INT, @qtd DECIMAL(16,3), @func INT = 0) AS
SET NOCOUNT ON
DECLARE @ant DECIMAL(16,3), @cod INT, @tipo NVARCHAR(20), @trans INT, @msg VARCHAR(200)

IF (@qtd IS NULL OR @qtd < 0)
 BEGIN
   RAISERROR('Quantidade contada inválida!', 16, 1)
   RETURN
 END

SET @func = ISNULL(@func, 0)
SET @trans = @@TRANCOUNT

BEGIN TRY
   IF (@trans = 0) BEGIN TRANSACTION

   --trava o produto para ler a quantidade atual (ADICAO/REMOCAO e calculada sobre o valor real no momento da gravacao)
   SELECT @ant = ISNULL(QUANT_ESTOQUE, 0) FROM Produtos WITH (UPDLOCK, HOLDLOCK) WHERE CODIGO = @produto AND ATIVO = 1
   IF (@@ROWCOUNT = 0)
    BEGIN
      RAISERROR('Produto não encontrado ou inativo!', 16, 1)
    END

   IF (@qtd <> @ant)
    BEGIN
      IF (@qtd > @ant)
         SET @tipo = N'ADI' + NCHAR(199) + NCHAR(195) + N'O'
      ELSE
         SET @tipo = N'REMO' + NCHAR(199) + NCHAR(195) + N'O'

      SELECT @cod = ISNULL(MAX(CODIGO), 0) + 1 FROM Produtos_Quant WITH (UPDLOCK, HOLDLOCK)

      INSERT INTO Produtos_Quant(CODIGO, COD_PRODUTO, COD_ENTRADA, FORMA, TIPO, QUANT, DATA, Hora, cod_usuario, Estoque)
      VALUES(@cod, @produto, 0, N'CONTAGEM', @tipo, ABS(@qtd - @ant),
             DATEADD(DAY, DATEDIFF(DAY, 0, GETDATE()), 0), CONVERT(DATETIME, CONVERT(VARCHAR(8), GETDATE(), 108)), @func, @qtd)
    END

   UPDATE Produtos SET QUANT_ESTOQUE = @qtd, CONTADO = 1, DATA_CONTAGEM = GETDATE() WHERE CODIGO = @produto

   IF (@trans = 0) COMMIT TRANSACTION
END TRY
BEGIN CATCH
   SET @msg = ERROR_MESSAGE()
   IF (@trans = 0 AND @@TRANCOUNT > 0) ROLLBACK TRANSACTION
   RAISERROR(@msg, 16, 1)
   RETURN
END CATCH

SELECT @ant AS QuantAnterior, @qtd AS QuantNova
GO

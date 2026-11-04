/***************************
PROJETO...: Online.info
ANALISTA..: Leandro Santos
DESCRICAO.: 
DATA......: 2022-04-23
MODIFICADO: 
***************************/
-- Idempotente: recria a function se ja existir. Usada pela view NotasDestinadas.
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF OBJECT_ID('dbo.SituacaoNF') IS NOT NULL DROP FUNCTION dbo.SituacaoNF
GO

CREATE FUNCTION dbo.SituacaoNF(@ChaveNF VARCHAR(44)) RETURNS VARCHAR(100) AS
BEGIN
   DECLARE @nfLancada INT
   
   SET @nfLancada = 0
   SELECT @nfLancada = ISNULL(COUNT(CodigoNota), 0) FROM EntradaEstoque WHERE ChavedeAcesso = @ChaveNF AND Enviada = 1
  	
	--Faz retorno
    RETURN @nfLancada
END
GO

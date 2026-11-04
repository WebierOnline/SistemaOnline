/***************************
PROJETO...: 
ANALISTA..: Victor Hugo Franchi Borges
DESCRICAO.: 
DATA......: 2011-11-15
ALTERACAO.: 
***************************/
-- Idempotente: recria a function se ja existir.
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF OBJECT_ID('dbo.hrStrMinutos') IS NOT NULL DROP FUNCTION dbo.hrStrMinutos
GO

CREATE FUNCTION dbo.hrStrMinutos(@hr VARCHAR(6)) RETURNS INT AS
BEGIN
 
	DECLARE @r INT, @p INT
	--Calcula
	SET @p = CHARINDEX(':', @hr)
	IF (@p > 0) SET @r = CONVERT(INT, RIGHT(@hr, LEN(@hr) - @p))
	--Faz retorno
	RETURN ISNULl(@r,0)

END
GO

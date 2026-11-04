/***************************
PROJETO...: 
ANALISTA..: Victor Hugo Franchi Borges
DESCRICAO.: 
DATA......: 2013-04-22
ALTERACAO.: 
***************************/
-- Idempotente: recria a function se ja existir.
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF OBJECT_ID('dbo.hrStrHoras') IS NOT NULL DROP FUNCTION dbo.hrStrHoras
GO

CREATE FUNCTION dbo.hrStrHoras(@hr VARCHAR(6)) RETURNS INT AS
BEGIN
 
	DECLARE @r INT, @p INT
	--Calcula
	SET @p = CHARINDEX(':', @hr)
	IF (@p > 0) SET @r = CONVERT(INT, SUBSTRING(@hr, 1, @p - 1))
	--Faz retorno
	RETURN ISNULl(@r,0)

END
GO

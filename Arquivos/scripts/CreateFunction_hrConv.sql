/***************************
PROJETO...: 
ANALISTA..: Evandro
DESCRICAO.: Converte uma hora em VARCHAR em DATETIME, exemplo: 05:35
DATA......: 2011-05-21
***************************/
-- Idempotente: recria a function se ja existir.
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF OBJECT_ID('dbo.hrConv') IS NOT NULL DROP FUNCTION dbo.hrConv
GO

CREATE FUNCTION dbo.hrConv(@hr VARCHAR(8), @dt DATETIME) RETURNS DATETIME AS
BEGIN
	DECLARE @dtConv VARCHAR(10)

	IF (NOT @dt IS NULL) SET @dtConv = CONVERT(char(10), @dt, 20)
	IF (@dt IS NULL) SET @dtConv = '1899-01-01'

	RETURN(CONVERT(DATETIME, @dtConv + ' ' + @hr))

END
GO

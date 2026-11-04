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
IF OBJECT_ID('dbo.hrStr') IS NOT NULL DROP FUNCTION dbo.hrStr
GO

CREATE FUNCTION dbo.hrStr() RETURNS VARCHAR(5) AS
BEGIN
 
	RETURN (CONVERT(char(5), GETDATE(), 108))

END
GO

/***************************
PROJETO...: 
ANALISTA..: Evandro
DESCRICAO.: 
DATA......: 2013-03-24
ALTERACAO.: 
***************************/
-- Idempotente: recria a function se ja existir.
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF OBJECT_ID('dbo.hrStr8') IS NOT NULL DROP FUNCTION dbo.hrStr8
GO

CREATE FUNCTION dbo.hrStr8() RETURNS VARCHAR(8) AS
BEGIN
 
	RETURN (CONVERT(char(8), GETDATE(), 108))

END
GO

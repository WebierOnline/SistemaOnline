/***************************
PROJETO...: SISC
ANALISTA..: Leandro Santos
DESCRICAO.: Função Retorna Data em formato String
DATA......: 2019-09-21
ALTERACAO.: 
***************************/
-- Idempotente: recria a function se ja existir.
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF OBJECT_ID('dbo.dtStr') IS NOT NULL DROP FUNCTION dbo.dtStr
GO

CREATE FUNCTION dbo.dtStr() RETURNS VARCHAR(10) AS
BEGIN
 
	RETURN (CONVERT(char(10), GETDATE(), 103))

END
GO

/***************************
PROJETO...: 
ANALISTA..: Victor Hugo Franchi Borges
DESCRICAO.: 
DATA......: 2010-10-14
ALTERACAO.: 
***************************/
-- Idempotente: recria a function se ja existir.
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF OBJECT_ID('dbo.hr') IS NOT NULL DROP FUNCTION dbo.hr
GO

CREATE FUNCTION dbo.hr() RETURNS DATETIME AS
BEGIN
 
 RETURN '1899-12-30 ' + (CONVERT(char(8), GETDATE(), 108))

END
GO

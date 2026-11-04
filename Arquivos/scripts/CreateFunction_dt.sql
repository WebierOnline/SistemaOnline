/***************************
PROJETO...: 
ANALISTA..: Victor Hugo Franchi Borges
DESCRICAO.: 
DATA......: 2010-05-25
ALTERACAO.: 
***************************/
-- Idempotente: recria a function se ja existir.
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF OBJECT_ID('dbo.dt') IS NOT NULL DROP FUNCTION dbo.dt
GO

CREATE FUNCTION dbo.dt() RETURNS DATETIME AS
BEGIN
 
 RETURN(CONVERT(char(10), GETDATE(), 20))

END
GO

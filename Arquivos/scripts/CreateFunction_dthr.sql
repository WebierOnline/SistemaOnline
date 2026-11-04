/***************************
PROJETO...: SISC
ANALISTA..: Leandro Santos
DESCRICAO.: Function que retorna Data e Hora
DATA......: 2016-02-26
ALTERACAO.: 
***************************/
-- Idempotente: recria a function se ja existir.
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF OBJECT_ID('dbo.dthr') IS NOT NULL DROP FUNCTION dbo.dthr
GO

CREATE FUNCTION dbo.dthr() RETURNS DATETIME AS
BEGIN
 
 RETURN(CONVERT(char(20), GETDATE(), 120))

END
GO

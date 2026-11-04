/***************************
PROJETO...: Online.Info
ANALISTA..: Leandro Santos
DESCRICAO.: View de Empresa usada para SPED
DATA......: 2022-03-14
ALTERACAO.: 
***************************/
-- Idempotente: cria um placeholder se o objeto ainda nao existir e sempre roda o ALTER com a definicao final.
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF OBJECT_ID('dbo.SPEDInventario') IS NULL
    EXEC('CREATE VIEW dbo.SPEDInventario AS SELECT 1 AS Placeholder')
GO

ALTER VIEW dbo.SPEDInventario AS

SELECT AnoInventario, PeriodoInventarioTipo, PeriodoInventario, Data, IdInventario, IdFilial 
FROM TbInventarios
GO

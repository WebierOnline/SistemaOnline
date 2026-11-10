/***************************
PROJETO...: Online.Info
DESCRICAO.: Atualizacao automatica das aliquotas CBS/IBS pela nuvem (arquivo gerado no Sistema_Financeiro).
            - Empresa.AliquotasReformaVersao: versao (AAAAMMDDHHMMSS) do ultimo arquivo aplicado.
            - Empresa.AliquotasReformaChecagem: data (AAAAMMDD) da ultima checagem na nuvem (1 vez por dia).
            - dbo.AplicarAliquotasReforma: quando varias vigencias cobrem a mesma UF/municipio no dia
              (ex.: regra "TODOS" e regra especifica de uma UF), vale a de inicio mais recente.
DATA......: 2026-11
***************************/
-- Idempotente.
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

IF COL_LENGTH('dbo.Empresa', 'AliquotasReformaVersao') IS NULL
    ALTER TABLE dbo.Empresa ADD AliquotasReformaVersao VARCHAR(14) NULL
GO

IF COL_LENGTH('dbo.Empresa', 'AliquotasReformaChecagem') IS NULL
    ALTER TABLE dbo.Empresa ADD AliquotasReformaChecagem VARCHAR(8) NULL
GO

IF OBJECT_ID('dbo.AplicarAliquotasReforma') IS NULL
    EXEC('CREATE PROCEDURE dbo.AplicarAliquotasReforma AS SELECT 1')
GO

ALTER PROCEDURE dbo.AplicarAliquotasReforma AS
SET NOCOUNT ON
DECLARE @hoje DATE, @cbs DECIMAL(10,4)
SET @hoje = CAST(GETDATE() AS DATE)

-- CBS (Empresa.CBSpAliq guarda 2 casas decimais)
SELECT TOP 1 @cbs = CBSpAliq
FROM CBS_Aliquota
WHERE @hoje BETWEEN dIniVig AND ISNULL(dFimVig, '99991231')
ORDER BY dIniVig DESC

IF @cbs IS NOT NULL
    IF EXISTS (SELECT 1 FROM Empresa WHERE CBSpAliq <> CAST(@cbs AS DECIMAL(10,2)))
        UPDATE Empresa SET CBSpAliq = CAST(@cbs AS DECIMAL(10,2))

-- IBS da UF (varias vigencias no mesmo dia: vale a de inicio mais recente)
;WITH V AS (
    SELECT E.IdEstado, E.IBSUFpAliq,
           ROW_NUMBER() OVER (PARTITION BY E.IdEstado ORDER BY E.dIniVig DESC, E.Id DESC) AS rn
    FROM IBS_Estado E
    WHERE @hoje BETWEEN E.dIniVig AND ISNULL(E.dFimVig, '99991231')
)
UPDATE C SET C.IBSUFpAliq = V.IBSUFpAliq
FROM Cidade C JOIN V ON C.IdEstado = V.IdEstado AND V.rn = 1
WHERE C.IBSUFpAliq <> V.IBSUFpAliq

-- IBS do municipio
;WITH V AS (
    SELECT M.CodigoMunicipio, M.IBSMunpAliq,
           ROW_NUMBER() OVER (PARTITION BY M.CodigoMunicipio ORDER BY M.dIniVig DESC, M.Id DESC) AS rn
    FROM IBS_Municipio M
    WHERE @hoje BETWEEN M.dIniVig AND ISNULL(M.dFimVig, '99991231')
)
UPDATE C SET C.IBSMunpAliq = V.IBSMunpAliq
FROM Cidade C JOIN V ON CAST(C.CodigoMunicipio AS NVARCHAR(7)) = V.CodigoMunicipio AND V.rn = 1
WHERE C.IBSMunpAliq <> V.IBSMunpAliq
GO

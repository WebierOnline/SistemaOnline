/***************************
PROJETO...: Online.Info
DESCRICAO.: Aliquotas da reforma tributaria (CBS e IBS) por VIGENCIA.
            - CBS_Aliquota: historico da aliquota da CBS (hoje era um valor unico em Empresa.CBSpAliq).
            - IBS 2027/2028: 0,05% UF + 0,05% municipio (art. 344 da LC 214/2025, NT 2025.002 regras UB18/UB37),
              inserido em IBS_Estado / IBS_Municipio com vigencia 01/01/2027 a 31/12/2028.
            - dbo.AplicarAliquotasReforma: copia as aliquotas vigentes HOJE para Empresa.CBSpAliq e Cidade.IBSUFpAliq /
              Cidade.IBSMunpAliq (as colunas que a NFe/NFCe ja leem). E chamada antes de calcular IBS/CBS de uma nota,
              entao a virada de ano e automatica. Sem linha de vigencia para hoje, mantem o valor atual.
            A CBS de 2027 NAO e inserida aqui: a NT manda usar a "Tabela de Aliquotas da CBS" oficial, ainda nao publicada.
            Quando sair, basta inserir a linha em CBS_Aliquota (dIniVig = 01/01/2027) e fechar a de 2026 (dFimVig = 31/12/2026).
DATA......: 2026-11
***************************/
-- Idempotente.
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- ---------------------------------------------------------------- CBS_Aliquota
IF OBJECT_ID('dbo.CBS_Aliquota') IS NULL
BEGIN
    CREATE TABLE dbo.CBS_Aliquota (
        Id        INT           IDENTITY(1,1) PRIMARY KEY,
        CBSpAliq  DECIMAL(10,4) NOT NULL,
        dIniVig   DATE          NOT NULL,
        dFimVig   DATE          NULL,
        Obs       NVARCHAR(200) NULL
    )
END
GO

IF NOT EXISTS (SELECT 1 FROM CBS_Aliquota)
    INSERT INTO CBS_Aliquota (CBSpAliq, dIniVig, dFimVig, Obs)
    VALUES (0.9, '20260101', '20261231', 'Ano-teste: pCBS deve ser 0,9% (art. 346 da LC 214/2025 - NT 2025.002 regra UB56-10)')
GO

-- ---------------------------------------------------------------- IBS 2027/2028 (0,05% UF + 0,05% municipio)
IF OBJECT_ID('dbo.IBS_Estado') IS NOT NULL
    INSERT INTO IBS_Estado (IdEstado, UF, IBSUFpAliq, dIniVig, dFimVig)
    SELECT DISTINCT e.IdEstado, e.UF, 0.05, '20270101', '20281231'
    FROM IBS_Estado e
    WHERE NOT EXISTS (SELECT 1 FROM IBS_Estado x WHERE x.IdEstado = e.IdEstado AND x.dIniVig = '20270101')
GO

IF OBJECT_ID('dbo.IBS_Municipio') IS NOT NULL
    INSERT INTO IBS_Municipio (CodigoMunicipio, IBSMunpAliq, dIniVig, dFimVig)
    SELECT DISTINCT m.CodigoMunicipio, 0.05, '20270101', '20281231'
    FROM IBS_Municipio m
    WHERE NOT EXISTS (SELECT 1 FROM IBS_Municipio x WHERE x.CodigoMunicipio = m.CodigoMunicipio AND x.dIniVig = '20270101')
GO

-- ---------------------------------------------------------------- procedure
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

-- IBS da UF
IF EXISTS (SELECT 1 FROM Cidade C JOIN IBS_Estado E ON C.IdEstado = E.IdEstado
           WHERE @hoje BETWEEN E.dIniVig AND ISNULL(E.dFimVig, '99991231') AND C.IBSUFpAliq <> E.IBSUFpAliq)
    UPDATE C SET C.IBSUFpAliq = E.IBSUFpAliq
    FROM Cidade C JOIN IBS_Estado E ON C.IdEstado = E.IdEstado
    WHERE @hoje BETWEEN E.dIniVig AND ISNULL(E.dFimVig, '99991231') AND C.IBSUFpAliq <> E.IBSUFpAliq

-- IBS do municipio
IF EXISTS (SELECT 1 FROM Cidade C JOIN IBS_Municipio M ON CAST(C.CodigoMunicipio AS NVARCHAR(7)) = M.CodigoMunicipio
           WHERE @hoje BETWEEN M.dIniVig AND ISNULL(M.dFimVig, '99991231') AND C.IBSMunpAliq <> M.IBSMunpAliq)
    UPDATE C SET C.IBSMunpAliq = M.IBSMunpAliq
    FROM Cidade C JOIN IBS_Municipio M ON CAST(C.CodigoMunicipio AS NVARCHAR(7)) = M.CodigoMunicipio
    WHERE @hoje BETWEEN M.dIniVig AND ISNULL(M.dFimVig, '99991231') AND C.IBSMunpAliq <> M.IBSMunpAliq
GO

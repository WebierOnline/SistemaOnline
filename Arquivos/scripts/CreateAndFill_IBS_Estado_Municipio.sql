-- ============================================================
-- Cria e preenche IBS_Estado e IBS_Municipio
-- Logica:
--   Cidade        = aliquota atual (consulta rapida em cadastros)
--   IBS_Estado    = historico de mudancas por estado
--   IBS_Municipio = historico de mudancas por municipio
-- ============================================================
SET NOCOUNT ON;
GO

-- ── IBS_Estado ────────────────────────────────────────────────
IF NOT EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID('IBS_Estado') AND type = 'U')
BEGIN
    CREATE TABLE IBS_Estado (
        Id          INT          IDENTITY(1,1) PRIMARY KEY,
        IdEstado    INT          NOT NULL,
        UF          NVARCHAR(2)  NOT NULL,
        IBSUFpAliq  DECIMAL(10,2) NOT NULL DEFAULT 0,
        dIniVig     DATE         NOT NULL,
        dFimVig     DATE         NULL
    );
END
GO

-- ── IBS_Municipio ─────────────────────────────────────────────
IF NOT EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID('IBS_Municipio') AND type = 'U')
BEGIN
    CREATE TABLE IBS_Municipio (
        Id               INT          IDENTITY(1,1) PRIMARY KEY,
        CodigoMunicipio  NVARCHAR(7)  NOT NULL,
        IBSMunpAliq      DECIMAL(10,2) NOT NULL DEFAULT 0,
        dIniVig          DATE         NOT NULL,
        dFimVig          DATE         NULL
    );
END
GO

-- ── Preenche IBS_Estado (1 registro por estado) ───────────────
IF NOT EXISTS (SELECT 1 FROM IBS_Estado)
BEGIN
    INSERT INTO IBS_Estado (IdEstado, UF, IBSUFpAliq, dIniVig, dFimVig)
    SELECT DISTINCT
        IdEstado,
        UF,
        0.10,
        CONVERT(date, '2026-01-01', 23),
        CONVERT(date, '2026-12-31', 23)
    FROM Cidade
    WHERE IdEstado IS NOT NULL
    ORDER BY UF;

    PRINT 'IBS_Estado: ' + CAST(@@ROWCOUNT AS VARCHAR) + ' estados inseridos.';
END
GO

-- ── Preenche IBS_Municipio (1 registro por municipio) ─────────
IF NOT EXISTS (SELECT 1 FROM IBS_Municipio)
BEGIN
    INSERT INTO IBS_Municipio (CodigoMunicipio, IBSMunpAliq, dIniVig, dFimVig)
    SELECT DISTINCT
        CAST(CodigoMunicipio AS NVARCHAR(7)),
        0.00,
        CONVERT(date, '2026-01-01', 23),
        CONVERT(date, '2026-12-31', 23)
    FROM Cidade
    WHERE CodigoMunicipio IS NOT NULL
    ORDER BY CAST(CodigoMunicipio AS NVARCHAR(7));

    PRINT 'IBS_Municipio: ' + CAST(@@ROWCOUNT AS VARCHAR) + ' municipios inseridos.';
END
GO

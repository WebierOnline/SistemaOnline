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
CREATE TABLE IBS_Estado (
    Id          INT          IDENTITY(1,1) PRIMARY KEY,
    IdEstado    INT          NOT NULL,
    UF          NVARCHAR(2)  NOT NULL,
    IBSUFpAliq  DECIMAL(10,2) NOT NULL DEFAULT 0,
    dIniVig     DATE         NOT NULL,
    dFimVig     DATE         NULL
);
GO

-- ── IBS_Municipio ─────────────────────────────────────────────
CREATE TABLE IBS_Municipio (
    Id               INT          IDENTITY(1,1) PRIMARY KEY,
    CodigoMunicipio  NVARCHAR(7)  NOT NULL,
    IBSMunpAliq      DECIMAL(10,2) NOT NULL DEFAULT 0,
    dIniVig          DATE         NOT NULL,
    dFimVig          DATE         NULL
);
GO

-- ── Preenche IBS_Estado (1 registro por estado) ───────────────
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
GO

PRINT 'IBS_Estado: ' + CAST(@@ROWCOUNT AS VARCHAR) + ' estados inseridos.';
GO

-- ── Preenche IBS_Municipio (1 registro por municipio) ─────────
INSERT INTO IBS_Municipio (CodigoMunicipio, IBSMunpAliq, dIniVig, dFimVig)
SELECT DISTINCT
    CAST(CodigoMunicipio AS NVARCHAR(7)),
    0.00,
    CONVERT(date, '2026-01-01', 23),
    CONVERT(date, '2026-12-31', 23)
FROM Cidade
WHERE CodigoMunicipio IS NOT NULL
ORDER BY CAST(CodigoMunicipio AS NVARCHAR(7));
GO

PRINT 'IBS_Municipio: ' + CAST(@@ROWCOUNT AS VARCHAR) + ' municipios inseridos.';
GO

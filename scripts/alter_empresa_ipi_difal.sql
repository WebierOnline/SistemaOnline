-- Controla se o IPI compoe a base de calculo do DIFAL
-- 0 = Nao inclui IPI na base do DIFAL (padrao conservador)
-- 1 = Inclui IPI na base do DIFAL (consumidor final interestadual)
--     conforme Art. 13, par. 1, inciso II da Lei Kandir
IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('Empresa') AND name = 'IPICompoeDIFAL'
)
    ALTER TABLE Empresa ADD IPICompoeDIFAL TINYINT NOT NULL DEFAULT 0;

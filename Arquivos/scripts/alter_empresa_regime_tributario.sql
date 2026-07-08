-- Regime tributario detalhado da empresa
-- 1 = Simples Nacional (CRT 1)
-- 2 = Simples Nacional - Excesso de Sublimite (CRT 2)
-- 3 = Lucro Presumido (CRT 3)
-- 4 = Lucro Real (CRT 3)
-- 5 = MEI (CRT 4)
IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('Empresa') AND name = 'RegimeTributario'
)
    ALTER TABLE Empresa ADD RegimeTributario TINYINT NULL;

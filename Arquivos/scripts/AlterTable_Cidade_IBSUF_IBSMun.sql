IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('Cidade') AND name = 'IBSUFpAliq'
)
    ALTER TABLE Cidade ADD [IBSUFpAliq]  DECIMAL(10,2) DEFAULT 0 NULL;
GO

IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('Cidade') AND name = 'IBSMunpAliq'
)
    ALTER TABLE Cidade ADD [IBSMunpAliq] DECIMAL(10,2) DEFAULT 0 NULL;
GO
PRINT 'Colunas IBSUFpAliq e IBSMunpAliq adicionadas.';
GO

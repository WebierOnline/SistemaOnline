ALTER TABLE Empresa
    ADD [CBSpAliq] DECIMAL(10,2) NOT NULL DEFAULT 0;
GO

UPDATE Empresa SET CBSpAliq = 0.90;
GO

PRINT 'CBSpAliq adicionado e preenchido: ' + CAST(@@ROWCOUNT AS VARCHAR) + ' registro(s).';
GO

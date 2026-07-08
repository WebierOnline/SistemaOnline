UPDATE Cidade
SET    IBSUFpAliq  = 0.10,
       IBSMunpAliq = 0.00
WHERE  IBSUFpAliq  IS NULL
   OR  IBSMunpAliq IS NULL;
GO

PRINT 'IBSUFpAliq / IBSMunpAliq atualizados: ' + CAST(@@ROWCOUNT AS VARCHAR) + ' registros.';
GO

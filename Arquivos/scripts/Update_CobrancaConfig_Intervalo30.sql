-- Cobranca por WhatsApp: intervalo minimo entre mensagens passa a ser 30 segundos (padrao 30 a 60).
-- Abaixo de 30 s o WhatsApp tende a bloquear o numero da empresa. O sistema tambem nao aceita menos que 30. Idempotente.
IF OBJECT_ID('dbo.CobrancaConfig', 'U') IS NOT NULL
BEGIN
    IF EXISTS (SELECT 1 FROM sys.default_constraints WHERE name = 'DF_CobrancaConfig_IntervaloMin')
        ALTER TABLE dbo.CobrancaConfig DROP CONSTRAINT DF_CobrancaConfig_IntervaloMin
    ALTER TABLE dbo.CobrancaConfig ADD CONSTRAINT DF_CobrancaConfig_IntervaloMin DEFAULT (30) FOR IntervaloMin
END
GO

IF OBJECT_ID('dbo.CobrancaConfig', 'U') IS NOT NULL
BEGIN
    UPDATE dbo.CobrancaConfig SET IntervaloMin = 30 WHERE IntervaloMin < 30
    UPDATE dbo.CobrancaConfig SET IntervaloMax = 60 WHERE IntervaloMax < IntervaloMin
END
GO

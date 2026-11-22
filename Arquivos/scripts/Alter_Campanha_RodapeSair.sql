-- Promocoes por WhatsApp: rodape Para nao receber mais promocoes, responda SAIR passa a ser opcional por campanha
-- (checkbox na tela, desmarcado por padrao). Idempotente.
IF OBJECT_ID('dbo.Campanha', 'U') IS NOT NULL AND COL_LENGTH('dbo.Campanha', 'RodapeSair') IS NULL
    ALTER TABLE dbo.Campanha ADD RodapeSair bit NOT NULL CONSTRAINT DF_Campanha_RodapeSair DEFAULT (0)
GO

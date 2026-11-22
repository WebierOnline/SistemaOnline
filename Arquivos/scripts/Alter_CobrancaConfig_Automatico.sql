-- Cobranca por WhatsApp - etapa 2 (envio automatico na maquina servidor: OnlineCommerce.exe /cobrancazap pela tarefa
-- agendada OnlineCommerce_CobrancaZap). CobrancaConfig.Automatico (script 184) liga/desliga; aqui as opcoes e a situacao:
--   AutoJuros          : cobrar multa + juros do atraso no automatico (padrao 1, como o checkbox da tela)
--   AutoResumoCelular  : celular que recebe o resumo de cada execucao (vazio = nao envia)
--   AutoUltimaExecucao / AutoUltimoResumo / AutoUltimoErro : ultima execucao (tela da Regua e alerta no popup)
-- Idempotente.
IF OBJECT_ID('dbo.CobrancaConfig', 'U') IS NOT NULL
BEGIN
    IF COL_LENGTH('dbo.CobrancaConfig', 'AutoJuros') IS NULL
        ALTER TABLE dbo.CobrancaConfig ADD AutoJuros bit NOT NULL CONSTRAINT DF_CobrancaConfig_AutoJuros DEFAULT (1)
    IF COL_LENGTH('dbo.CobrancaConfig', 'AutoResumoCelular') IS NULL
        ALTER TABLE dbo.CobrancaConfig ADD AutoResumoCelular nvarchar(20) NULL
    IF COL_LENGTH('dbo.CobrancaConfig', 'AutoUltimaExecucao') IS NULL
        ALTER TABLE dbo.CobrancaConfig ADD AutoUltimaExecucao datetime NULL
    IF COL_LENGTH('dbo.CobrancaConfig', 'AutoUltimoResumo') IS NULL
        ALTER TABLE dbo.CobrancaConfig ADD AutoUltimoResumo nvarchar(300) NULL
    IF COL_LENGTH('dbo.CobrancaConfig', 'AutoUltimoErro') IS NULL
        ALTER TABLE dbo.CobrancaConfig ADD AutoUltimoErro nvarchar(300) NULL
END
GO

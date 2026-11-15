-- NFS-e: forma dos tributos aproximados (Lei 12.741, grupo totTrib da DPS) escolhida no Configurar NFS-e,
-- igual as opcoes do portal nacional (Perfil > Configuracao > Valor aproximado dos tributos):
--   NFSeFormaTotTrib: 0 automatico pelo regime (ME/EPP = % do Simples; fora do Simples = percentuais),
--                     1 valores em R$ (calculados em cada nota pelos % federal/estadual/municipal),
--                     2 percentuais federal/estadual/municipal, 3 aliquota do Simples Nacional (so ME/EPP).
--   MEI: sempre "nao informar" (indTotTrib), qualquer que seja a forma.
--   NFSeTotTribEst: % estadual (as colunas NFSeTotTribSN/Fed/Mun sao do script 178). Idempotente.
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

IF COL_LENGTH('dbo.empresa', 'NFSeFormaTotTrib') IS NULL
    ALTER TABLE dbo.empresa ADD NFSeFormaTotTrib tinyint NOT NULL CONSTRAINT DF_empresa_NFSeFormaTotTrib DEFAULT (0)
GO
IF COL_LENGTH('dbo.empresa', 'NFSeTotTribEst') IS NULL
    ALTER TABLE dbo.empresa ADD NFSeTotTribEst decimal(5, 2) NOT NULL CONSTRAINT DF_empresa_NFSeTotTribEst DEFAULT (0)
GO

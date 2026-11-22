-- Cobranca por WhatsApp: opcao por regra de listar na mensagem tambem as parcelas que ainda vao vencer
-- (frmCobrancaRegua, chkParcelasEmDias). Padrao 1 = lista, inclusive nas regras que ja existem. Idempotente.
IF OBJECT_ID('dbo.CobrancaRegra', 'U') IS NOT NULL AND COL_LENGTH('dbo.CobrancaRegra', 'ListarAVencer') IS NULL
    ALTER TABLE dbo.CobrancaRegra ADD ListarAVencer bit NOT NULL CONSTRAINT DF_CobrancaRegra_ListarAVencer DEFAULT (1)
GO

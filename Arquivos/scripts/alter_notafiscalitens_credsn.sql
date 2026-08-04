-- Adiciona campos de Crédito Simples Nacional (CSOSN 101/201) na tabela NotaFiscalItens
-- pCredSN  = alíquota de crédito do emitente (vem de Empresa.pCreditoICMSSimplesNacional)
-- vCredICMSSN = valor do crédito calculado (vBC * pCredSN / 100)
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('NotaFiscalItens') AND name = 'pCredSN')
    ALTER TABLE NotaFiscalItens ADD pCredSN DECIMAL(7,4) NOT NULL DEFAULT 0;
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('NotaFiscalItens') AND name = 'vCredICMSSN')
    ALTER TABLE NotaFiscalItens ADD vCredICMSSN DECIMAL(15,2) NOT NULL DEFAULT 0;
GO

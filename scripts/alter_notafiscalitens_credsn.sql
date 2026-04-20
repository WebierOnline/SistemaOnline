-- Adiciona campos de Crédito Simples Nacional (CSOSN 101/201) na tabela NotaFiscalItens
-- pCredSN  = alíquota de crédito do emitente (vem de Empresa.pCreditoICMSSimplesNacional)
-- vCredICMSSN = valor do crédito calculado (vBC * pCredSN / 100)
ALTER TABLE NotaFiscalItens
    ADD pCredSN     DECIMAL(7,4)  NOT NULL DEFAULT 0,
        vCredICMSSN DECIMAL(15,2) NOT NULL DEFAULT 0;

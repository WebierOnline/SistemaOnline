-- Adiciona campo uTrib_IS à tabela NotaFiscalItens existente
-- Preenchido automaticamente a partir de tbISClassTrib ao adicionar item

ALTER TABLE [dbo].[NotaFiscalItens]
    ADD [uTrib_IS] VARCHAR(6) NULL;

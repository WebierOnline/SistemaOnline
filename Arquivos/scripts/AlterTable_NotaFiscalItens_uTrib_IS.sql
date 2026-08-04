-- Adiciona campo uTrib_IS à tabela NotaFiscalItens existente
-- Preenchido automaticamente a partir de tbISClassTrib ao adicionar item

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('NotaFiscalItens') AND name = 'uTrib_IS')
    ALTER TABLE [dbo].[NotaFiscalItens]
    ADD [uTrib_IS] VARCHAR(6) NULL;

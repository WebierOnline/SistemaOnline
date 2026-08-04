-- Adiciona campo uTrib_IS à tabela tbISClassTrib existente
-- Armazena a unidade tributável do IS (ex: 'UN', 'ML', 'KG')

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('tbISClassTrib') AND name = 'uTrib_IS')
    ALTER TABLE [dbo].[tbISClassTrib]
    ADD [uTrib_IS] VARCHAR(6) NULL;
GO


-- Popula 'UN' em todos os registros existentes
UPDATE [dbo].[tbISClassTrib]
    SET [uTrib_IS] = 'UN'
    WHERE [uTrib_IS] IS NULL;

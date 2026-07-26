-- Adiciona campo uTrib_IS à tabela tbISClassTrib existente
-- Armazena a unidade tributável do IS (ex: 'UN', 'ML', 'KG')

ALTER TABLE [dbo].[tbISClassTrib]
    ADD [uTrib_IS] VARCHAR(6) NULL;

-- Popula 'UN' em todos os registros existentes
UPDATE [dbo].[tbISClassTrib]
    SET [uTrib_IS] = 'UN'
    WHERE [uTrib_IS] IS NULL;

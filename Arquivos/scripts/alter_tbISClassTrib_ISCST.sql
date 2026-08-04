-- Adiciona campo ISCST em tbISClassTrib
-- ISCST: CST do Imposto Seletivo (00/01/99) referente à classificação IS

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('tbISClassTrib') AND name = 'ISCST')
    ALTER TABLE [dbo].[tbISClassTrib]
    ADD [ISCST] VARCHAR(2) NULL;
GO


-- Todos os registros atuais são "Saída Tributada (Comércio/Varejo)"
UPDATE [dbo].[tbISClassTrib] SET [ISCST] = '01';

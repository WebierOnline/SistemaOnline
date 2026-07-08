-- Adiciona campo ISCST em tbISClassTrib
-- ISCST: CST do Imposto Seletivo (00/01/99) referente à classificação IS

ALTER TABLE [dbo].[tbISClassTrib]
    ADD [ISCST] VARCHAR(2) NULL;

-- Todos os registros atuais são "Saída Tributada (Comércio/Varejo)"
UPDATE [dbo].[tbISClassTrib] SET [ISCST] = '01';

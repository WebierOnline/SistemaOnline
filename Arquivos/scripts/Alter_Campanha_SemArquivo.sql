-- Promocoes por WhatsApp: o arquivo da campanha (imagem, video ou PDF) deixa de ficar no banco.
-- Agora fica em <pasta do sistema>\Campanhas\<nome da campanha> no servidor (Campanha.ArquivoNome = caminho relativo)
-- e e apagado ao cancelar ou excluir a campanha. Remove a coluna Campanha.Arquivo (varbinary). Idempotente.
IF OBJECT_ID('dbo.Campanha', 'U') IS NOT NULL AND COL_LENGTH('dbo.Campanha', 'Arquivo') IS NOT NULL
    ALTER TABLE dbo.Campanha DROP COLUMN Arquivo
GO

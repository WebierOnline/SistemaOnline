-- Promocoes por WhatsApp: contatos importados do celular (Online Commerce > Promocoes por WhatsApp > Importar contatos).
--   Contatos.ATIVO = 0: o contato aparece em cinza na lista e nao recebe (botao Desativar / Ativar contato).
--   CampanhaDestinatario.CodContato: destinatario que e contato do celular (CodCliente = 0). Idempotente.
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF OBJECT_ID('dbo.Contatos', 'U') IS NULL
BEGIN
    CREATE TABLE dbo.Contatos (
        Id            int IDENTITY(1,1) NOT NULL CONSTRAINT PK_Contatos PRIMARY KEY,
        Nome          nvarchar(160)  NOT NULL,
        Celular       nvarchar(20)   NOT NULL,
        ATIVO         bit            NOT NULL CONSTRAINT DF_Contatos_ATIVO DEFAULT (1),
        ImportadoEm   datetime       NOT NULL CONSTRAINT DF_Contatos_ImportadoEm DEFAULT (GETDATE()),
        AtualizadoEm  datetime       NULL
    )
    CREATE UNIQUE INDEX UX_Contatos_Celular ON dbo.Contatos (Celular)
END
GO
IF OBJECT_ID('dbo.CampanhaDestinatario', 'U') IS NOT NULL AND COL_LENGTH('dbo.CampanhaDestinatario', 'CodContato') IS NULL
    ALTER TABLE dbo.CampanhaDestinatario ADD CodContato int NULL
GO
IF OBJECT_ID('dbo.CampanhaDestinatario', 'U') IS NOT NULL
   AND NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_CampanhaDestinatario_Celular' AND object_id = OBJECT_ID('dbo.CampanhaDestinatario'))
    CREATE INDEX IX_CampanhaDestinatario_Celular ON dbo.CampanhaDestinatario (Celular, DataHora)
GO

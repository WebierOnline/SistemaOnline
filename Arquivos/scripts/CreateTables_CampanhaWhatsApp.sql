-- Promocoes e avisos por WhatsApp (Online Commerce > Financeiro > Promocoes por WhatsApp). Idempotente.
--   Campanha             : uma campanha (texto, midia opcional: imagem, video ou PDF guardado no banco, agendamento e situacao).
--                          TipoMidia 0 = so texto, 1 = texto + imagem, 2 = texto + video, 3 = so imagem, 4 = texto + PDF.
--                          Situacao 1 = na fila (enviando), 2 = pausada, 3 = concluida, 4 = cancelada.
--   CampanhaDestinatario : clientes da campanha. Situacao 0 = aguardando, 1 = enviada, 2 = erro, 3 = pulada.
--   cliente.NaoReceberPromocao : cliente que nao quer promocoes (a cobranca continua).
--   CobrancaConfig.PromoLimiteDia : promocoes por dia (padrao 40; a cobranca tem limite proprio e prioridade).
-- Quem envia e a tarefa do servidor (OnlineCommerce.exe /cobrancazap), no horario configurado na Regua de cobranca.
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF OBJECT_ID('dbo.Campanha', 'U') IS NULL
BEGIN
    CREATE TABLE dbo.Campanha (
        Id            int IDENTITY(1,1) NOT NULL CONSTRAINT PK_Campanha PRIMARY KEY,
        Nome          nvarchar(100)  NOT NULL,
        Mensagem      nvarchar(2000) NULL,
        TipoMidia     tinyint        NOT NULL CONSTRAINT DF_Campanha_TipoMidia DEFAULT (0),
        ArquivoNome   nvarchar(200)  NULL,
        Arquivo       varbinary(max) NULL,
        Publico       nvarchar(1000) NULL,
        AgendadaPara  datetime       NULL,
        Situacao      tinyint        NOT NULL CONSTRAINT DF_Campanha_Situacao DEFAULT (1),
        CriadaEm      datetime       NOT NULL CONSTRAINT DF_Campanha_CriadaEm DEFAULT (GETDATE()),
        Usuario       nvarchar(50)   NULL,
        ConcluidaEm   datetime       NULL
    )
END
GO
IF OBJECT_ID('dbo.CampanhaDestinatario', 'U') IS NULL
BEGIN
    CREATE TABLE dbo.CampanhaDestinatario (
        Id           int IDENTITY(1,1) NOT NULL CONSTRAINT PK_CampanhaDestinatario PRIMARY KEY,
        CodCampanha  int            NOT NULL,
        CodCliente   int            NOT NULL,
        Celular      nvarchar(20)   NULL,
        Situacao     tinyint        NOT NULL CONSTRAINT DF_CampanhaDestinatario_Situacao DEFAULT (0),
        DataHora     datetime       NULL,
        Erro         nvarchar(300)  NULL
    )
    CREATE INDEX IX_CampanhaDestinatario_Campanha ON dbo.CampanhaDestinatario (CodCampanha, Situacao)
    CREATE INDEX IX_CampanhaDestinatario_Cliente ON dbo.CampanhaDestinatario (CodCliente, DataHora)
END
GO
IF COL_LENGTH('dbo.cliente', 'NaoReceberPromocao') IS NULL
    ALTER TABLE dbo.cliente ADD NaoReceberPromocao bit NOT NULL CONSTRAINT DF_cliente_NaoReceberPromocao DEFAULT (0)
GO
IF OBJECT_ID('dbo.CobrancaConfig', 'U') IS NOT NULL AND COL_LENGTH('dbo.CobrancaConfig', 'PromoLimiteDia') IS NULL
    ALTER TABLE dbo.CobrancaConfig ADD PromoLimiteDia int NOT NULL CONSTRAINT DF_CobrancaConfig_PromoLimiteDia DEFAULT (40)
GO

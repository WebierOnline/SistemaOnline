-- Cobranca por WhatsApp (Online Commerce, Financeiro > Cobranca por WhatsApp).
--   CobrancaConfig       : linha unica (Id = 1) com horario, limite por dia, intervalo entre mensagens e chave Pix.
--   CobrancaRegra        : regua de cobranca. Situacao 1 = antes do vencimento, 2 = no dia, 3 = apos o vencimento;
--                          Dias = quantos dias antes/depois; RepetirDias 0 = nao repete; MaxEnvios por parcela.
--   CobrancaEnvio        : historico (uma linha por mensagem: um cliente pode ter varias parcelas na mesma mensagem).
--                          Situacao 1 = enviada, 2 = erro. Manual 1 = tela, 0 = automatico.
--   CobrancaEnvioParcela : parcelas que dispararam a regra naquela mensagem (controle de repeticao).
--   cliente.NaoCobrarWhatsApp : cliente que nao deve receber cobranca pelo WhatsApp.
-- As regras padrao sao criadas pelo proprio sistema na primeira vez que a tela abre. Idempotente.
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

IF OBJECT_ID('dbo.CobrancaConfig', 'U') IS NULL
BEGIN
    CREATE TABLE dbo.CobrancaConfig (
        Id             int            NOT NULL CONSTRAINT PK_CobrancaConfig PRIMARY KEY,
        HoraInicio     varchar(5)     NOT NULL CONSTRAINT DF_CobrancaConfig_HoraInicio DEFAULT ('08:00'),
        HoraFim        varchar(5)     NOT NULL CONSTRAINT DF_CobrancaConfig_HoraFim DEFAULT ('18:00'),
        LimiteDia      int            NOT NULL CONSTRAINT DF_CobrancaConfig_LimiteDia DEFAULT (50),
        IntervaloMin   int            NOT NULL CONSTRAINT DF_CobrancaConfig_IntervaloMin DEFAULT (20),
        IntervaloMax   int            NOT NULL CONSTRAINT DF_CobrancaConfig_IntervaloMax DEFAULT (60),
        EnviarSabado   bit            NOT NULL CONSTRAINT DF_CobrancaConfig_EnviarSabado DEFAULT (0),
        EnviarDomingo  bit            NOT NULL CONSTRAINT DF_CobrancaConfig_EnviarDomingo DEFAULT (0),
        ChavePix       nvarchar(100)  NULL,
        Automatico     bit            NOT NULL CONSTRAINT DF_CobrancaConfig_Automatico DEFAULT (0)
    )
END
GO

IF NOT EXISTS (SELECT 1 FROM dbo.CobrancaConfig WHERE Id = 1)
    INSERT INTO dbo.CobrancaConfig (Id) VALUES (1)
GO

IF OBJECT_ID('dbo.CobrancaRegra', 'U') IS NULL
BEGIN
    CREATE TABLE dbo.CobrancaRegra (
        Id           int IDENTITY(1,1) NOT NULL CONSTRAINT PK_CobrancaRegra PRIMARY KEY,
        Descricao    nvarchar(60)    NOT NULL,
        Ativa        bit             NOT NULL CONSTRAINT DF_CobrancaRegra_Ativa DEFAULT (1),
        Situacao     tinyint         NOT NULL,
        Dias         int             NOT NULL CONSTRAINT DF_CobrancaRegra_Dias DEFAULT (0),
        RepetirDias  int             NOT NULL CONSTRAINT DF_CobrancaRegra_RepetirDias DEFAULT (0),
        MaxEnvios    int             NOT NULL CONSTRAINT DF_CobrancaRegra_MaxEnvios DEFAULT (1),
        Mensagem     nvarchar(2000)  NOT NULL
    )
END
GO

IF OBJECT_ID('dbo.CobrancaEnvio', 'U') IS NULL
BEGIN
    CREATE TABLE dbo.CobrancaEnvio (
        Id          int IDENTITY(1,1) NOT NULL CONSTRAINT PK_CobrancaEnvio PRIMARY KEY,
        DataHora    datetime        NOT NULL CONSTRAINT DF_CobrancaEnvio_DataHora DEFAULT (GETDATE()),
        CodCliente  int             NOT NULL,
        CodRegra    int             NULL,
        Celular     varchar(20)     NULL,
        ValorTotal  decimal(15,2)   NOT NULL CONSTRAINT DF_CobrancaEnvio_ValorTotal DEFAULT (0),
        Situacao    tinyint         NOT NULL,
        Erro        nvarchar(300)   NULL,
        Manual      bit             NOT NULL CONSTRAINT DF_CobrancaEnvio_Manual DEFAULT (1),
        Usuario     nvarchar(50)    NULL
    )
    CREATE INDEX IX_CobrancaEnvio_Cliente ON dbo.CobrancaEnvio (CodCliente, DataHora)
    CREATE INDEX IX_CobrancaEnvio_DataHora ON dbo.CobrancaEnvio (DataHora)
END
GO

IF OBJECT_ID('dbo.CobrancaEnvioParcela', 'U') IS NULL
BEGIN
    CREATE TABLE dbo.CobrancaEnvioParcela (
        CodEnvio    int  NOT NULL,
        CodParcela  int  NOT NULL,
        CodRegra    int  NOT NULL,
        CONSTRAINT PK_CobrancaEnvioParcela PRIMARY KEY (CodEnvio, CodParcela)
    )
    CREATE INDEX IX_CobrancaEnvioParcela_Parcela ON dbo.CobrancaEnvioParcela (CodParcela, CodRegra)
END
GO

IF COL_LENGTH('dbo.cliente', 'NaoCobrarWhatsApp') IS NULL
    ALTER TABLE dbo.cliente ADD NaoCobrarWhatsApp bit NULL
GO

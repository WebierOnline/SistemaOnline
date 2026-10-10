-- Tabelas da licenca online (servidor de licencas: bloqueio, Pix "Pagar agora" e codigo de liberacao).
-- Mesmas tabelas do Online Food (schema od). Usadas pelo modulo Compartilhado\Modulos\LicencaOnline.bas
-- (OnlineCommerce, PDV e Ordem de Servico), que tambem cria tudo isto sozinho se faltar.
-- Idempotente: so cria o que ainda nao existir.

IF SCHEMA_ID('od') IS NULL EXEC('CREATE SCHEMA od');

IF OBJECT_ID('od.Configuracoes') IS NULL
    CREATE TABLE od.Configuracoes (
        Chave varchar(50)   NOT NULL CONSTRAINT PK_Configuracoes PRIMARY KEY,
        Valor nvarchar(200) NULL
    );

IF OBJECT_ID('od.Licenca') IS NULL
    CREATE TABLE od.Licenca (
        Id                  tinyint      NOT NULL CONSTRAINT PK_Licenca PRIMARY KEY CONSTRAINT CK_Licenca_Unica CHECK (Id = 1),
        Token               varchar(max) NULL,
        AtualizadoEm        datetime     NOT NULL CONSTRAINT DF_Licenca_Atualizado DEFAULT (GETDATE()),
        UltimaVerificacao   datetime     NULL,
        UltimaDataVista     datetime     NULL,
        LiberadoAte         datetime     NULL,
        TemporarioUsadoPara date         NULL
    );

IF OBJECT_ID('od.LicencaCodigosUsados') IS NULL
    CREATE TABLE od.LicencaCodigosUsados (
        Codigo     char(8)  NOT NULL CONSTRAINT PK_LicencaCodigosUsados PRIMARY KEY,
        Tipo       tinyint  NOT NULL,
        UsadoEm    datetime NOT NULL CONSTRAINT DF_LicCod_UsadoEm DEFAULT (GETDATE()),
        CodUsuario int      NULL
    );

IF NOT EXISTS (SELECT 1 FROM od.Licenca WHERE Id = 1)
    INSERT INTO od.Licenca (Id) VALUES (1);

-- Cria tabela TabelaIBPT para importacao do arquivo IBPT
IF OBJECT_ID('TabelaIBPT', 'U') IS NULL
BEGIN
    CREATE TABLE TabelaIBPT (
        codigo            VARCHAR(10)    NOT NULL,
        ex                VARCHAR(2)     NOT NULL DEFAULT '0',
        tipo              TINYINT        NOT NULL DEFAULT 0,
        descricao         VARCHAR(500)   NOT NULL DEFAULT '',
        nacionalfederal   DECIMAL(10,4)  NOT NULL DEFAULT 0,
        importadosfederal DECIMAL(10,4)  NOT NULL DEFAULT 0,
        estadual          DECIMAL(10,4)  NOT NULL DEFAULT 0,
        municipal         DECIMAL(10,4)  NOT NULL DEFAULT 0,
        vigenciainicio    DATE           NULL,
        vigenciafim       DATE           NULL,
        chave             VARCHAR(10)    NOT NULL DEFAULT '',
        versao            VARCHAR(10)    NOT NULL DEFAULT '',
        fonte             VARCHAR(60)    NOT NULL DEFAULT '',
        CONSTRAINT PK_TabelaIBPT PRIMARY KEY (codigo, ex)
    );
    PRINT 'TabelaIBPT criada com sucesso.';
END
ELSE
    PRINT 'TabelaIBPT ja existe.';
GO

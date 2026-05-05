-- Tabela tbNCM: configuracoes por NCM (aliquotas + reforma tributaria)
-- Aliquotas (nacionalfederal/importadosfederal/estadual/municipal) sao sincronizadas
-- automaticamente a cada importacao do TabelaIBPT.
-- Os campos de reforma (cClassTrib_IBS, cClassTrib_IS, tipo_calculo_is)
-- sao gerenciados manualmente pelo usuario.
IF OBJECT_ID('tbNCM', 'U') IS NULL
BEGIN
    CREATE TABLE tbNCM (
        NCM                 VARCHAR(10)    NOT NULL,
        descricao           VARCHAR(500)   NOT NULL DEFAULT '',
        nacionalfederal     DECIMAL(10,4)  NOT NULL DEFAULT 0,
        importadosfederal   DECIMAL(10,4)  NOT NULL DEFAULT 0,
        estadual            DECIMAL(10,4)  NOT NULL DEFAULT 0,
        municipal           DECIMAL(10,4)  NOT NULL DEFAULT 0,
        cClassTrib_IBS      VARCHAR(6)     NOT NULL DEFAULT '',
        cClassTrib_IS       VARCHAR(6)     NOT NULL DEFAULT '',
        tipo_calculo_is     INTEGER        NOT NULL DEFAULT 0,
        CONSTRAINT PK_tbNCM PRIMARY KEY (NCM)
    );
    PRINT 'tbNCM criada com sucesso.';
END
ELSE
    PRINT 'tbNCM ja existe.';
GO

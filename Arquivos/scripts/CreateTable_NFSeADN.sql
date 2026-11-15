-- NFS-e: documentos da distribuicao do ADN (Ambiente de Dados Nacional), GET /contribuintes/DFe/{NSU}.
-- O NFSeEmissor (acao=distribuir) baixa em lotes, a partir do maior NSU gravado, tudo em que a empresa aparece:
-- NFS-e emitidas (inclusive as feitas direto no portal), RECEBIDAS (empresa como tomador ou intermediario) e eventos.
--   Papel: 1 emitente, 2 tomador, 3 intermediario, 0 outro.  Situacao: 1 normal, 3 cancelada (evento de cancelamento
--   recebido para a chave).  IdNFSe: nota do sistema com a mesma chave (emitida pelo OrdemServico).
-- O XML fica no banco (o PDF e gerado sob demanda pela acao=pdfadn). Idempotente.
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

IF OBJECT_ID('dbo.NFSeADN', 'U') IS NULL
BEGIN
    CREATE TABLE dbo.NFSeADN (
        Ambiente         tinyint        NOT NULL,
        NSU              bigint         NOT NULL,
        ChaveAcesso      varchar(50)    NULL,
        TipoDocumento    varchar(30)    NOT NULL,
        TipoEvento       varchar(6)     NULL,
        Papel            tinyint        NOT NULL CONSTRAINT DF_NFSeADN_Papel DEFAULT (0),
        Situacao         tinyint        NOT NULL CONSTRAINT DF_NFSeADN_Situacao DEFAULT (1),
        NumeroNFSe       varchar(15)    NULL,
        DataEmissao      datetime       NULL,
        EmitenteDoc      varchar(14)    NULL,
        EmitenteNome     varchar(300)   NULL,
        TomadorDoc       varchar(14)    NULL,
        TomadorNome      varchar(300)   NULL,
        ValorServico     decimal(15,2)  NULL,
        ValorLiquido     decimal(15,2)  NULL,
        CodTribNac       varchar(6)     NULL,
        Descricao        varchar(2000)  NULL,
        Municipio        varchar(150)   NULL,
        Xml              nvarchar(max)  NULL,
        IdNFSe           int            NULL,
        DataRecebimento  datetime       NOT NULL CONSTRAINT DF_NFSeADN_DataRecebimento DEFAULT (GETDATE()),
        CONSTRAINT PK_NFSeADN PRIMARY KEY (Ambiente, NSU)
    )
    CREATE INDEX IX_NFSeADN_Chave ON dbo.NFSeADN (ChaveAcesso)
    CREATE INDEX IX_NFSeADN_Papel ON dbo.NFSeADN (Papel, DataEmissao)
END
GO

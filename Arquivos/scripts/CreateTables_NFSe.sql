-- NFS-e Nacional (API direta do Sefin) a partir das OS finalizadas do OrdemServico.
-- Emissao feita pelo NFSeEmissor.exe (C:\ProjetoNET\NFSeNacional), chamado pela tela do OS.
-- empresa: parametros da NFS-e (ambiente, serie/ultimo numero da DPS, codigos padrao do servico, regime, tributos).
-- NFSe: uma nota por linha (copia dos dados do tomador e dos valores no momento da geracao; o que vai ao Sefin
--       e o que esta aqui, entao a tela pode corrigir antes de transmitir).
--   Situacao: 0 Em digitacao, 1 Autorizada, 2 Rejeitada, 3 Cancelada, 4 Enviando/Pendente (falha de comunicacao:
--   consultar antes de reenviar).
-- NFSeItens: servicos das OS que formam a nota (uma OS so pode estar em uma nota ativa).
-- View NFSe_OSDisponiveis: OS finalizadas com servico e sem NFS-e ativa. Idempotente.
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- ---------------------------------------------------------------- empresa
IF COL_LENGTH('dbo.empresa', 'NFSeAmbiente') IS NULL
    ALTER TABLE dbo.empresa ADD NFSeAmbiente tinyint NOT NULL CONSTRAINT DF_empresa_NFSeAmbiente DEFAULT (2)
GO
IF COL_LENGTH('dbo.empresa', 'NFSeSerieDPS') IS NULL
    ALTER TABLE dbo.empresa ADD NFSeSerieDPS int NOT NULL CONSTRAINT DF_empresa_NFSeSerieDPS DEFAULT (1)
GO
IF COL_LENGTH('dbo.empresa', 'NFSeUltimoDPS') IS NULL
    ALTER TABLE dbo.empresa ADD NFSeUltimoDPS bigint NOT NULL CONSTRAINT DF_empresa_NFSeUltimoDPS DEFAULT (0)
GO
IF COL_LENGTH('dbo.empresa', 'NFSeCodTribNac') IS NULL
    ALTER TABLE dbo.empresa ADD NFSeCodTribNac varchar(6) NULL
GO
IF COL_LENGTH('dbo.empresa', 'NFSeCodTribMun') IS NULL
    ALTER TABLE dbo.empresa ADD NFSeCodTribMun varchar(3) NULL
GO
IF COL_LENGTH('dbo.empresa', 'NFSeCodNBS') IS NULL
    ALTER TABLE dbo.empresa ADD NFSeCodNBS varchar(9) NULL
GO
-- regEspTrib (0 nenhum ... 9 outros) e regApTribSN (1 tudo pelo SN, 2 ISS fora, 3 tudo fora; so ME/EPP)
IF COL_LENGTH('dbo.empresa', 'NFSeRegimeEspecial') IS NULL
    ALTER TABLE dbo.empresa ADD NFSeRegimeEspecial tinyint NOT NULL CONSTRAINT DF_empresa_NFSeRegimeEspecial DEFAULT (0)
GO
IF COL_LENGTH('dbo.empresa', 'NFSeRegimeApuracaoSN') IS NULL
    ALTER TABLE dbo.empresa ADD NFSeRegimeApuracaoSN tinyint NOT NULL CONSTRAINT DF_empresa_NFSeRegimeApuracaoSN DEFAULT (1)
GO
-- aliquota do ISS: so vai na DPS quando a regra exige (ex.: ISS retido pelo tomador no Simples)
IF COL_LENGTH('dbo.empresa', 'NFSeAliquotaISS') IS NULL
    ALTER TABLE dbo.empresa ADD NFSeAliquotaISS decimal(5, 2) NULL
GO
-- totais aproximados de tributos (Lei 12.741): % do Simples (ME/EPP) ou % federal/municipal (nao optante)
IF COL_LENGTH('dbo.empresa', 'NFSeTotTribSN') IS NULL
    ALTER TABLE dbo.empresa ADD NFSeTotTribSN decimal(5, 2) NOT NULL CONSTRAINT DF_empresa_NFSeTotTribSN DEFAULT (0)
GO
IF COL_LENGTH('dbo.empresa', 'NFSeTotTribFed') IS NULL
    ALTER TABLE dbo.empresa ADD NFSeTotTribFed decimal(5, 2) NOT NULL CONSTRAINT DF_empresa_NFSeTotTribFed DEFAULT (0)
GO
IF COL_LENGTH('dbo.empresa', 'NFSeTotTribMun') IS NULL
    ALTER TABLE dbo.empresa ADD NFSeTotTribMun decimal(5, 2) NOT NULL CONSTRAINT DF_empresa_NFSeTotTribMun DEFAULT (0)
GO
IF COL_LENGTH('dbo.empresa', 'NFSeMensagem') IS NULL
    ALTER TABLE dbo.empresa ADD NFSeMensagem nvarchar(500) NULL
GO
-- grupo IBS/CBS (opcional em 2026; exige NBS)
IF COL_LENGTH('dbo.empresa', 'NFSeIBSCBS') IS NULL
    ALTER TABLE dbo.empresa ADD NFSeIBSCBS bit NOT NULL CONSTRAINT DF_empresa_NFSeIBSCBS DEFAULT (0)
GO
IF COL_LENGTH('dbo.empresa', 'NFSeIndOp') IS NULL
    ALTER TABLE dbo.empresa ADD NFSeIndOp varchar(6) NULL
GO
IF COL_LENGTH('dbo.empresa', 'NFSeCST') IS NULL
    ALTER TABLE dbo.empresa ADD NFSeCST varchar(3) NULL
GO
IF COL_LENGTH('dbo.empresa', 'NFSeClassTrib') IS NULL
    ALTER TABLE dbo.empresa ADD NFSeClassTrib varchar(6) NULL
GO
-- assinatura: 0 = RSA-SHA1 (perfil NF-e), 1 = RSA-SHA256
IF COL_LENGTH('dbo.empresa', 'NFSeAssinaturaSHA256') IS NULL
    ALTER TABLE dbo.empresa ADD NFSeAssinaturaSHA256 bit NOT NULL CONSTRAINT DF_empresa_NFSeAssinaturaSHA256 DEFAULT (0)
GO

-- ---------------------------------------------------------------- NFSe
IF OBJECT_ID('dbo.NFSe', 'U') IS NULL
BEGIN
    CREATE TABLE dbo.NFSe (
        IdNFSe                  int IDENTITY(1,1) NOT NULL CONSTRAINT PK_NFSe PRIMARY KEY,
        Situacao                tinyint        NOT NULL CONSTRAINT DF_NFSe_Situacao DEFAULT (0),
        Ambiente                tinyint        NOT NULL CONSTRAINT DF_NFSe_Ambiente DEFAULT (2),
        SerieDPS                int            NULL,
        NumeroDPS               bigint         NULL,
        IdDPS                   varchar(45)    NULL,
        DataEmissao             datetime       NULL,
        DataCompetencia         datetime       NOT NULL,
        CodCliente              int            NULL,
        TomadorCpfCnpj          varchar(14)    NULL,
        TomadorNome             nvarchar(150)  NULL,
        TomadorIM               varchar(15)    NULL,
        TomadorLogradouro       nvarchar(255)  NULL,
        TomadorNumero           nvarchar(60)   NULL,
        TomadorComplemento      nvarchar(156)  NULL,
        TomadorBairro           nvarchar(60)   NULL,
        TomadorCodMunicipio     int            NULL,
        TomadorCEP              varchar(8)     NULL,
        TomadorFone             varchar(20)    NULL,
        TomadorEmail            nvarchar(80)   NULL,
        CodTribNac              varchar(6)     NOT NULL,
        CodTribMun              varchar(3)     NULL,
        CodNBS                  varchar(9)     NULL,
        CodMunicipioPrestacao   int            NOT NULL,
        Descricao               nvarchar(2000) NOT NULL,
        InfoComplementar        nvarchar(2000) NULL,
        NumeroPedido            nvarchar(60)   NULL,
        ValorServicos           decimal(15, 2) NOT NULL CONSTRAINT DF_NFSe_ValorServicos DEFAULT (0),
        DescontoIncondicionado  decimal(15, 2) NOT NULL CONSTRAINT DF_NFSe_DescontoIncond DEFAULT (0),
        DescontoCondicionado    decimal(15, 2) NOT NULL CONSTRAINT DF_NFSe_DescontoCond DEFAULT (0),
        RetencaoISS             tinyint        NOT NULL CONSTRAINT DF_NFSe_RetencaoISS DEFAULT (1),
        AliquotaISS             decimal(5, 2)  NULL,
        -- devolvidos pelo Sefin
        BaseCalculo             decimal(15, 2) NULL,
        AliquotaAplicada        decimal(5, 2)  NULL,
        ValorISS                decimal(15, 2) NULL,
        ValorLiquido            decimal(15, 2) NULL,
        ChaveAcesso             varchar(50)    NULL,
        NumeroNFSe              varchar(13)    NULL,
        DataAutorizacao         datetime       NULL,
        XmlNFSe                 nvarchar(max)  NULL,
        ArquivoXml              nvarchar(260)  NULL,
        ArquivoPdf              nvarchar(260)  NULL,
        Mensagem                nvarchar(2000) NULL,
        -- cancelamento
        DataCancelamento        datetime       NULL,
        MotivoCancelamento      nvarchar(255)  NULL,
        XmlCancelamento         nvarchar(max)  NULL,
        Usuario                 nvarchar(50)   NULL,
        DataCadastro            datetime       NOT NULL CONSTRAINT DF_NFSe_DataCadastro DEFAULT (GETDATE())
    )
    CREATE INDEX IX_NFSe_Situacao ON dbo.NFSe (Situacao, DataCompetencia)
    CREATE INDEX IX_NFSe_Chave ON dbo.NFSe (ChaveAcesso)
    CREATE INDEX IX_NFSe_Cliente ON dbo.NFSe (CodCliente)
END
GO

IF OBJECT_ID('dbo.NFSeItens', 'U') IS NULL
BEGIN
    CREATE TABLE dbo.NFSeItens (
        IdNFSeItem      int IDENTITY(1,1) NOT NULL CONSTRAINT PK_NFSeItens PRIMARY KEY,
        IdNFSe          int            NOT NULL CONSTRAINT FK_NFSeItens_NFSe REFERENCES dbo.NFSe (IdNFSe) ON DELETE CASCADE,
        CodOS           int            NOT NULL,
        CodServicoOS    int            NULL,   -- OS_Servicos_Auto.codigo
        Descricao       nvarchar(120)  NOT NULL,
        Quantidade      decimal(15, 4) NOT NULL CONSTRAINT DF_NFSeItens_Quantidade DEFAULT (1),
        ValorUnitario   decimal(15, 4) NOT NULL CONSTRAINT DF_NFSeItens_ValorUnitario DEFAULT (0),
        Desconto        decimal(15, 2) NOT NULL CONSTRAINT DF_NFSeItens_Desconto DEFAULT (0),
        Total           decimal(15, 2) NOT NULL CONSTRAINT DF_NFSeItens_Total DEFAULT (0)
    )
    CREATE INDEX IX_NFSeItens_NFSe ON dbo.NFSeItens (IdNFSe)
    CREATE INDEX IX_NFSeItens_OS ON dbo.NFSeItens (CodOS)
END
GO

-- ---------------------------------------------------------------- OS disponiveis para NFS-e
IF OBJECT_ID('dbo.NFSe_OSDisponiveis', 'V') IS NOT NULL
    DROP VIEW dbo.NFSe_OSDisponiveis
GO
CREATE VIEW dbo.NFSe_OSDisponiveis
AS
SELECT o.COD_OS, o.COD_CLIENTE, c.Nome AS Cliente, c.CPF AS CpfCnpj, o.DATA_TERMINO, o.TIPO_OS,
       s.ValorServicos, s.DescontoServicos
FROM dbo.OS o
INNER JOIN (SELECT cod_os,
                   SUM(ISNULL(preco, 0) * ISNULL(quantidade, 0)) AS ValorServicos,
                   SUM(ISNULL(desconto, 0)) AS DescontoServicos
            FROM dbo.OS_Servicos_Auto
            GROUP BY cod_os) s ON s.cod_os = o.COD_OS
LEFT JOIN dbo.cliente c ON c.CODIGO = o.COD_CLIENTE
WHERE ISNULL(o.STATUS_OS, 0) = 1
  AND s.ValorServicos > 0
  AND NOT EXISTS (SELECT 1 FROM dbo.NFSeItens i INNER JOIN dbo.NFSe n ON n.IdNFSe = i.IdNFSe
                  WHERE i.CodOS = o.COD_OS AND n.Situacao IN (0, 1, 2, 4))
GO

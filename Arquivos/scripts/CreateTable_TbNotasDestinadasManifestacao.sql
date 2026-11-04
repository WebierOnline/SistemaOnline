-- TbNotasDestinadasManifestacao: notas destinadas ao CNPJ (manifestacao do destinatario).
-- Idempotente: cria a tabela se nao existir e acrescenta so as colunas que faltarem
-- (cliente sem nada, com a tabela completa ou com uma versao antiga/parcial).
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO

IF OBJECT_ID('dbo.TbNotasDestinadasManifestacao', 'U') IS NULL
BEGIN
CREATE TABLE [dbo].[TbNotasDestinadasManifestacao](
	[CodigoConsulta] [int] NOT NULL,
	[Data] [datetime] NULL,
	[TipoRetornoWS] [varchar](15) NOT NULL,
	[TipoNotaWS] [smallint] NOT NULL,
	[CNPJFornecedor] [varchar](18) NOT NULL,
	[RazaoSocial] [varchar](60) NOT NULL,
	[ChaveAcesso] [varchar](44) NOT NULL,
	[DataEmissao] [datetime] NULL,
	[ValorNota] [decimal](11, 2) NOT NULL,
	[SituacaoNota] [smallint] NOT NULL,
	[SituacaoConfirmacao] [smallint] NOT NULL,
	[nSeqEvento] [smallint] NOT NULL,
	[TextoCorrecao] [text] NOT NULL,
	[TipoEvento] [varchar](38) NOT NULL,
	[Justificativa] [varchar](255) NOT NULL,
	[NumeroProtocolo] [bigint] NOT NULL,
	[DataHoraProcotolo] [varchar](25) NOT NULL,
	[Enviada] [bit] NOT NULL,
	[Enviar] [bit] NOT NULL,
	[XMLAutorizado] [text] NOT NULL,
	[ImportadoXML] [bit] NOT NULL,
	[IdFilial] [smallint] NOT NULL,
	[IdNFConsulta] [bigint] IDENTITY(1,1) NOT FOR REPLICATION NOT NULL,
	[DataInclusao] [datetime] NULL,
	[TipoEventoCodigo] [varchar](8) NOT NULL,
	[EventoProtocolo] [bigint] NOT NULL,
	[EventoDataHora] [varchar](25) NOT NULL,
	[quick~rs] [timestamp] NULL,
 CONSTRAINT [IdNFConsulta] PRIMARY KEY CLUSTERED 
(
	[IdNFConsulta] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON, FILLFACTOR = 90) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.TbNotasDestinadasManifestacao') AND name = 'CodigoConsulta')
    ALTER TABLE dbo.TbNotasDestinadasManifestacao ADD [CodigoConsulta] [int] NOT NULL DEFAULT (0)
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.TbNotasDestinadasManifestacao') AND name = 'TipoRetornoWS')
    ALTER TABLE dbo.TbNotasDestinadasManifestacao ADD [TipoRetornoWS] [varchar](15) NOT NULL DEFAULT ('')
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.TbNotasDestinadasManifestacao') AND name = 'TipoNotaWS')
    ALTER TABLE dbo.TbNotasDestinadasManifestacao ADD [TipoNotaWS] [smallint] NOT NULL DEFAULT (0)
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.TbNotasDestinadasManifestacao') AND name = 'CNPJFornecedor')
    ALTER TABLE dbo.TbNotasDestinadasManifestacao ADD [CNPJFornecedor] [varchar](18) NOT NULL DEFAULT ('')
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.TbNotasDestinadasManifestacao') AND name = 'RazaoSocial')
    ALTER TABLE dbo.TbNotasDestinadasManifestacao ADD [RazaoSocial] [varchar](60) NOT NULL DEFAULT ('')
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.TbNotasDestinadasManifestacao') AND name = 'ChaveAcesso')
    ALTER TABLE dbo.TbNotasDestinadasManifestacao ADD [ChaveAcesso] [varchar](44) NOT NULL DEFAULT ('')
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.TbNotasDestinadasManifestacao') AND name = 'ValorNota')
    ALTER TABLE dbo.TbNotasDestinadasManifestacao ADD [ValorNota] [decimal](11, 2) NOT NULL DEFAULT (0)
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.TbNotasDestinadasManifestacao') AND name = 'SituacaoNota')
    ALTER TABLE dbo.TbNotasDestinadasManifestacao ADD [SituacaoNota] [smallint] NOT NULL DEFAULT (0)
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.TbNotasDestinadasManifestacao') AND name = 'SituacaoConfirmacao')
    ALTER TABLE dbo.TbNotasDestinadasManifestacao ADD [SituacaoConfirmacao] [smallint] NOT NULL DEFAULT (0)
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.TbNotasDestinadasManifestacao') AND name = 'nSeqEvento')
    ALTER TABLE dbo.TbNotasDestinadasManifestacao ADD [nSeqEvento] [smallint] NOT NULL DEFAULT (0)
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.TbNotasDestinadasManifestacao') AND name = 'TextoCorrecao')
    ALTER TABLE dbo.TbNotasDestinadasManifestacao ADD [TextoCorrecao] [text] NOT NULL DEFAULT ('')
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.TbNotasDestinadasManifestacao') AND name = 'TipoEvento')
    ALTER TABLE dbo.TbNotasDestinadasManifestacao ADD [TipoEvento] [varchar](38) NOT NULL DEFAULT ('')
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.TbNotasDestinadasManifestacao') AND name = 'Justificativa')
    ALTER TABLE dbo.TbNotasDestinadasManifestacao ADD [Justificativa] [varchar](255) NOT NULL DEFAULT ('')
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.TbNotasDestinadasManifestacao') AND name = 'NumeroProtocolo')
    ALTER TABLE dbo.TbNotasDestinadasManifestacao ADD [NumeroProtocolo] [bigint] NOT NULL DEFAULT (0)
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.TbNotasDestinadasManifestacao') AND name = 'DataHoraProcotolo')
    ALTER TABLE dbo.TbNotasDestinadasManifestacao ADD [DataHoraProcotolo] [varchar](25) NOT NULL DEFAULT ('')
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.TbNotasDestinadasManifestacao') AND name = 'Enviada')
    ALTER TABLE dbo.TbNotasDestinadasManifestacao ADD [Enviada] [bit] NOT NULL DEFAULT (0)
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.TbNotasDestinadasManifestacao') AND name = 'Enviar')
    ALTER TABLE dbo.TbNotasDestinadasManifestacao ADD [Enviar] [bit] NOT NULL DEFAULT (0)
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.TbNotasDestinadasManifestacao') AND name = 'XMLAutorizado')
    ALTER TABLE dbo.TbNotasDestinadasManifestacao ADD [XMLAutorizado] [text] NOT NULL DEFAULT ('')
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.TbNotasDestinadasManifestacao') AND name = 'ImportadoXML')
    ALTER TABLE dbo.TbNotasDestinadasManifestacao ADD [ImportadoXML] [bit] NOT NULL DEFAULT (0)
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.TbNotasDestinadasManifestacao') AND name = 'IdFilial')
    ALTER TABLE dbo.TbNotasDestinadasManifestacao ADD [IdFilial] [smallint] NOT NULL DEFAULT (0)
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.TbNotasDestinadasManifestacao') AND name = 'TipoEventoCodigo')
    ALTER TABLE dbo.TbNotasDestinadasManifestacao ADD [TipoEventoCodigo] [varchar](8) NOT NULL DEFAULT ('')
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.TbNotasDestinadasManifestacao') AND name = 'EventoProtocolo')
    ALTER TABLE dbo.TbNotasDestinadasManifestacao ADD [EventoProtocolo] [bigint] NOT NULL DEFAULT (0)
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.TbNotasDestinadasManifestacao') AND name = 'EventoDataHora')
    ALTER TABLE dbo.TbNotasDestinadasManifestacao ADD [EventoDataHora] [varchar](25) NOT NULL DEFAULT ('')
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.TbNotasDestinadasManifestacao') AND name = 'Data')
    ALTER TABLE dbo.TbNotasDestinadasManifestacao ADD [Data] [datetime] NULL
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.TbNotasDestinadasManifestacao') AND name = 'DataEmissao')
    ALTER TABLE dbo.TbNotasDestinadasManifestacao ADD [DataEmissao] [datetime] NULL
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.TbNotasDestinadasManifestacao') AND name = 'DataInclusao')
    ALTER TABLE dbo.TbNotasDestinadasManifestacao ADD [DataInclusao] [datetime] NULL
GO

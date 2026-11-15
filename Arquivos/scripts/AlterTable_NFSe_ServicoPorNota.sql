-- NFS-e: dados do servico escolhidos em cada nota no Editar NFS-e (antes vinham sempre do Configurar):
--   IndOp: indicador da operacao (cIndOp do grupo IBS/CBS, Anexo VII) - ex.: 050101 servico feito no estabelecimento,
--          050102 feito em outro local (na casa/empresa do cliente). NULL = o do Configurar (empresa.NFSeIndOp).
--   ClassTrib: cClassTrib da nota (o CST sao os 3 primeiros digitos). NULL = o do Configurar.
--   PrestacaoNoCliente: 1 = servico feito fora do estabelecimento; o municipio fica em CodMunicipioPrestacao (script 178).
-- Idempotente.
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

IF COL_LENGTH('dbo.NFSe', 'IndOp') IS NULL
    ALTER TABLE dbo.NFSe ADD IndOp varchar(6) NULL
GO
IF COL_LENGTH('dbo.NFSe', 'ClassTrib') IS NULL
    ALTER TABLE dbo.NFSe ADD ClassTrib varchar(6) NULL
GO
IF COL_LENGTH('dbo.NFSe', 'PrestacaoNoCliente') IS NULL
    ALTER TABLE dbo.NFSe ADD PrestacaoNoCliente bit NOT NULL CONSTRAINT DF_NFSe_PrestacaoNoCliente DEFAULT (0)
GO

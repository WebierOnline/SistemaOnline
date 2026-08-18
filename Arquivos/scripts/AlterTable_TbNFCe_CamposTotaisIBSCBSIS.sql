-- Totais de IBS/CBS/IS no cabecalho da NFCe (TbNFCe), para relatorio de apuracao
-- fiscal sem precisar somar TbNFCe_Itens toda vez. Espelha o mesmo padrao ja usado
-- na NotaFiscal da NFe (NFe_Completa.frm: SUM(IBS_vBC)/SUM(CBS_vBC) por nota) -
-- vBCIBS e vBCCBS separados (na pratica iguais, pois IBS_vBC e CBS_vBC do item tem
-- a mesma base), nao um campo combinado.
-- Nao exigido pela SEFAZ (o documento valido e o XML assinado), e conveniencia
-- de relatorio/auditoria contabil.
-- Todas NULL sem DEFAULT (metadata-only, seguro em SQL Server pre-2012/Express).

IF EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('TbNFCe') AND name = 'vBCIBSCBS')
    ALTER TABLE TbNFCe DROP COLUMN vBCIBSCBS;
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('TbNFCe') AND name = 'vBCCBS')
    ALTER TABLE TbNFCe ADD vBCCBS DECIMAL(15, 2) NULL;
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('TbNFCe') AND name = 'vBCIBS')
    ALTER TABLE TbNFCe ADD vBCIBS DECIMAL(15, 2) NULL;
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('TbNFCe') AND name = 'vIBSUF')
    ALTER TABLE TbNFCe ADD vIBSUF DECIMAL(15, 2) NULL;
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('TbNFCe') AND name = 'vIBSMun')
    ALTER TABLE TbNFCe ADD vIBSMun DECIMAL(15, 2) NULL;
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('TbNFCe') AND name = 'vIBS')
    ALTER TABLE TbNFCe ADD vIBS DECIMAL(15, 2) NULL;
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('TbNFCe') AND name = 'vCBS')
    ALTER TABLE TbNFCe ADD vCBS DECIMAL(15, 2) NULL;
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('TbNFCe') AND name = 'vBCIS')
    ALTER TABLE TbNFCe ADD vBCIS DECIMAL(15, 2) NULL;
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('TbNFCe') AND name = 'vIS')
    ALTER TABLE TbNFCe ADD vIS DECIMAL(15, 2) NULL;
GO

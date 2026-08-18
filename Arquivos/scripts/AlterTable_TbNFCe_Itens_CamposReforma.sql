-- Campos da reforma tributaria (IBS/CBS/IS) em TbNFCe_Itens, espelhando os ja existentes em
-- NotaFiscalItens (AlterTable_NotaFiscalItens_CamposReforma.sql) para o PDV/NFCe tambem poder emitir com reforma.
-- Todas NULL sem DEFAULT (diferente do script da NotaFiscalItens): "ADD COLUMN NOT NULL DEFAULT" reescreve
-- a tabela inteira em SQL Server pre-2012/Express, e TbNFCe_Itens pode ter volume grande em cliente antigo.
-- O codigo VB6 que le essas colunas ja trata NULL via IIf(IsNull(...), 0, ...) em tudo.

-- [Bloco IBS]
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('TbNFCe_Itens') AND name = 'cClassTrib')
    ALTER TABLE TbNFCe_Itens ADD cClassTrib VARCHAR(20) NULL;
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('TbNFCe_Itens') AND name = 'IBSCBS_CST')
    ALTER TABLE TbNFCe_Itens ADD IBSCBS_CST VARCHAR(3) NULL;
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('TbNFCe_Itens') AND name = 'IBS_vBC')
    ALTER TABLE TbNFCe_Itens ADD IBS_vBC DECIMAL(15, 2) NULL;
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('TbNFCe_Itens') AND name = 'IBS_UFpAliq')
    ALTER TABLE TbNFCe_Itens ADD IBS_UFpAliq DECIMAL(5, 2) NULL;
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('TbNFCe_Itens') AND name = 'IBS_MunpAliq')
    ALTER TABLE TbNFCe_Itens ADD IBS_MunpAliq DECIMAL(5, 2) NULL;
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('TbNFCe_Itens') AND name = 'IBS_pRed')
    ALTER TABLE TbNFCe_Itens ADD IBS_pRed DECIMAL(5, 2) NULL;
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('TbNFCe_Itens') AND name = 'IBS_vIBSUF')
    ALTER TABLE TbNFCe_Itens ADD IBS_vIBSUF DECIMAL(15, 2) NULL;
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('TbNFCe_Itens') AND name = 'IBS_vIBSMun')
    ALTER TABLE TbNFCe_Itens ADD IBS_vIBSMun DECIMAL(15, 2) NULL;
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('TbNFCe_Itens') AND name = 'IBS_vIBS')
    ALTER TABLE TbNFCe_Itens ADD IBS_vIBS DECIMAL(15, 2) NULL;
GO

-- [Bloco CBS]
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('TbNFCe_Itens') AND name = 'CBS_vBC')
    ALTER TABLE TbNFCe_Itens ADD CBS_vBC DECIMAL(15, 2) NULL;
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('TbNFCe_Itens') AND name = 'CBS_pAliq')
    ALTER TABLE TbNFCe_Itens ADD CBS_pAliq DECIMAL(5, 2) NULL;
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('TbNFCe_Itens') AND name = 'CBS_pRed')
    ALTER TABLE TbNFCe_Itens ADD CBS_pRed DECIMAL(5, 2) NULL;
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('TbNFCe_Itens') AND name = 'CBS_vCBS')
    ALTER TABLE TbNFCe_Itens ADD CBS_vCBS DECIMAL(15, 2) NULL;
GO

-- [Bloco IS]
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('TbNFCe_Itens') AND name = 'cClassTrib_IS')
    ALTER TABLE TbNFCe_Itens ADD cClassTrib_IS VARCHAR(20) NULL;
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('TbNFCe_Itens') AND name = 'IS_CST')
    ALTER TABLE TbNFCe_Itens ADD IS_CST VARCHAR(2) NULL;
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('TbNFCe_Itens') AND name = 'IS_tipo_calculo')
    ALTER TABLE TbNFCe_Itens ADD IS_tipo_calculo INT NULL; -- 0=sem IS, 1=% (Ad Valorem), 2=Fixo (Especifica), 3=Mista
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('TbNFCe_Itens') AND name = 'IS_vBC')
    ALTER TABLE TbNFCe_Itens ADD IS_vBC DECIMAL(15, 2) NULL;
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('TbNFCe_Itens') AND name = 'IS_pAliq')
    ALTER TABLE TbNFCe_Itens ADD IS_pAliq DECIMAL(5, 2) NULL;
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('TbNFCe_Itens') AND name = 'IS_qUnid')
    ALTER TABLE TbNFCe_Itens ADD IS_qUnid DECIMAL(15, 4) NULL;
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('TbNFCe_Itens') AND name = 'IS_vUnid')
    ALTER TABLE TbNFCe_Itens ADD IS_vUnid DECIMAL(15, 4) NULL;
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('TbNFCe_Itens') AND name = 'IS_vIS')
    ALTER TABLE TbNFCe_Itens ADD IS_vIS DECIMAL(15, 2) NULL;
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('TbNFCe_Itens') AND name = 'uTrib_IS')
    ALTER TABLE TbNFCe_Itens ADD uTrib_IS VARCHAR(6) NULL;
GO

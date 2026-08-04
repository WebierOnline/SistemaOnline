-- Rodar no SQL Server Management Studio (SSMS) para atualizar a tabela NotaFiscalItens

-- [Bloco IBS]
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('NotaFiscalItens') AND name = 'cClassTrib')
    ALTER TABLE NotaFiscalItens ADD cClassTrib VARCHAR(20) NULL;
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('NotaFiscalItens') AND name = 'IBSCBS_CST')
    ALTER TABLE NotaFiscalItens ADD IBSCBS_CST VARCHAR(3) NULL;
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('NotaFiscalItens') AND name = 'IBS_vBC')
    ALTER TABLE NotaFiscalItens ADD IBS_vBC DECIMAL(15, 2) NOT NULL DEFAULT 0;
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('NotaFiscalItens') AND name = 'IBS_UFpAliq')
    ALTER TABLE NotaFiscalItens ADD IBS_UFpAliq DECIMAL(5, 2) NOT NULL DEFAULT 0;
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('NotaFiscalItens') AND name = 'IBS_MunpAliq')
    ALTER TABLE NotaFiscalItens ADD IBS_MunpAliq DECIMAL(5, 2) NOT NULL DEFAULT 0;
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('NotaFiscalItens') AND name = 'IBS_pRed')
    ALTER TABLE NotaFiscalItens ADD IBS_pRed DECIMAL(5, 2) NOT NULL DEFAULT 0;
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('NotaFiscalItens') AND name = 'IBS_vIBSUF')
    ALTER TABLE NotaFiscalItens ADD IBS_vIBSUF DECIMAL(15, 2) NOT NULL DEFAULT 0;
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('NotaFiscalItens') AND name = 'IBS_vIBSMun')
    ALTER TABLE NotaFiscalItens ADD IBS_vIBSMun DECIMAL(15, 2) NOT NULL DEFAULT 0;
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('NotaFiscalItens') AND name = 'IBS_vIBS')
    ALTER TABLE NotaFiscalItens ADD IBS_vIBS DECIMAL(15, 2) NOT NULL DEFAULT 0;
GO

-- [Bloco CBS]
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('NotaFiscalItens') AND name = 'CBS_vBC')
    ALTER TABLE NotaFiscalItens ADD CBS_vBC DECIMAL(15, 2) NOT NULL DEFAULT 0; -- Adicionado para segurança fiscal
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('NotaFiscalItens') AND name = 'CBS_pAliq')
    ALTER TABLE NotaFiscalItens ADD CBS_pAliq DECIMAL(5, 2) NOT NULL DEFAULT 0;
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('NotaFiscalItens') AND name = 'CBS_pRed')
    ALTER TABLE NotaFiscalItens ADD CBS_pRed DECIMAL(5, 2) NOT NULL DEFAULT 0;
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('NotaFiscalItens') AND name = 'CBS_vCBS')
    ALTER TABLE NotaFiscalItens ADD CBS_vCBS DECIMAL(15, 2) NOT NULL DEFAULT 0;
GO

-- [Bloco IS]
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('NotaFiscalItens') AND name = 'cClassTrib_IS')
    ALTER TABLE NotaFiscalItens ADD cClassTrib_IS VARCHAR(20) NULL;
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('NotaFiscalItens') AND name = 'IS_CST')
    ALTER TABLE NotaFiscalItens ADD IS_CST VARCHAR(2) NULL;
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('NotaFiscalItens') AND name = 'IS_tipo_calculo')
    ALTER TABLE NotaFiscalItens ADD IS_tipo_calculo INT NOT NULL DEFAULT 1; -- 1=% (Ad Valorem), 2=Fixo (Específica)
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('NotaFiscalItens') AND name = 'IS_vBC')
    ALTER TABLE NotaFiscalItens ADD IS_vBC DECIMAL(15, 2) NOT NULL DEFAULT 0;
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('NotaFiscalItens') AND name = 'IS_pAliq')
    ALTER TABLE NotaFiscalItens ADD IS_pAliq DECIMAL(5, 2) NOT NULL DEFAULT 0;
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('NotaFiscalItens') AND name = 'IS_qUnid')
    ALTER TABLE NotaFiscalItens ADD IS_qUnid DECIMAL(15, 4) NOT NULL DEFAULT 0; -- 4 casas decimais para litragem/unidades
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('NotaFiscalItens') AND name = 'IS_vUnid')
    ALTER TABLE NotaFiscalItens ADD IS_vUnid DECIMAL(15, 4) NOT NULL DEFAULT 0; -- 4 casas decimais para valor por unidade
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('NotaFiscalItens') AND name = 'IS_vIS')
    ALTER TABLE NotaFiscalItens ADD IS_vIS DECIMAL(15, 2) NOT NULL DEFAULT 0;
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('NotaFiscalItens') AND name = 'uTrib_IS')
    ALTER TABLE NotaFiscalItens ADD uTrib_IS VARCHAR(6) NULL;
GO
-- Adiciona totais IBS / CBS / IS em NotaFiscal
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('NotaFiscal') AND name = 'vBCCBS')
    ALTER TABLE NotaFiscal ADD vBCCBS DECIMAL(15,2) NULL;
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('NotaFiscal') AND name = 'vBCIBS')
    ALTER TABLE NotaFiscal ADD vBCIBS DECIMAL(15,2) NULL;
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('NotaFiscal') AND name = 'vIBSUF')
    ALTER TABLE NotaFiscal ADD vIBSUF DECIMAL(15,2) NULL;
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('NotaFiscal') AND name = 'vIBSMun')
    ALTER TABLE NotaFiscal ADD vIBSMun DECIMAL(15,2) NULL;
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('NotaFiscal') AND name = 'vIBS')
    ALTER TABLE NotaFiscal ADD vIBS DECIMAL(15,2) NULL;
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('NotaFiscal') AND name = 'vCBS')
    ALTER TABLE NotaFiscal ADD vCBS DECIMAL(15,2) NULL;
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('NotaFiscal') AND name = 'vBCIS')
    ALTER TABLE NotaFiscal ADD vBCIS DECIMAL(15,2) NULL;
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('NotaFiscal') AND name = 'vIS')
    ALTER TABLE NotaFiscal ADD vIS DECIMAL(15,2) NULL;
GO

-- Migracao: se a coluna vBCCBSIBS ja existir, renomear / recriar
-- EXEC sp_rename 'NotaFiscal.vBCCBSIBS', 'vBCIBS', 'COLUMN'
-- ALTER TABLE NotaFiscal ADD vBCCBS DECIMAL(15,2) NULL
-- UPDATE NotaFiscal SET vBCCBS = vBCIBS

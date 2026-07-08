-- Adiciona totais IBS / CBS / IS em NotaFiscal
ALTER TABLE NotaFiscal ADD
    vBCCBS     DECIMAL(15,2) NULL,
    vBCIBS     DECIMAL(15,2) NULL,
    vIBSUF     DECIMAL(15,2) NULL,
    vIBSMun    DECIMAL(15,2) NULL,
    vIBS       DECIMAL(15,2) NULL,
    vCBS       DECIMAL(15,2) NULL,
    vBCIS      DECIMAL(15,2) NULL,
    vIS        DECIMAL(15,2) NULL
GO

-- Migracao: se a coluna vBCCBSIBS ja existir, renomear / recriar
-- EXEC sp_rename 'NotaFiscal.vBCCBSIBS', 'vBCIBS', 'COLUMN'
-- ALTER TABLE NotaFiscal ADD vBCCBS DECIMAL(15,2) NULL
-- UPDATE NotaFiscal SET vBCCBS = vBCIBS

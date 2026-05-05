-- Adiciona totais IBS / CBS / IS em NotaFiscal
ALTER TABLE NotaFiscal ADD
    vBCCBSIBS  DECIMAL(15,2) NULL,
    vIBSUF     DECIMAL(15,2) NULL,
    vIBSMun    DECIMAL(15,2) NULL,
    vIBS       DECIMAL(15,2) NULL,
    vCBS       DECIMAL(15,2) NULL,
    vBCIS      DECIMAL(15,2) NULL,
    vIS        DECIMAL(15,2) NULL
GO

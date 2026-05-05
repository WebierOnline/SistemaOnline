-- Adiciona campos IBS / CBS / IS em NotaFiscalItens
ALTER TABLE NotaFiscalItens ADD
    IBSCBSCST   VARCHAR(3)    NULL,
    CBSpAliq    DECIMAL(8,4)  NULL,
    IBSUFpAliq  DECIMAL(8,4)  NULL,
    IBSMunpAliq DECIMAL(8,4)  NULL,
    vBCCBSIBS   DECIMAL(15,2) NULL,
    vIBSUF      DECIMAL(15,2) NULL,
    vIBSMun     DECIMAL(15,2) NULL,
    vIBS        DECIMAL(15,2) NULL,
    vCBS        DECIMAL(15,2) NULL,
    ISCST       VARCHAR(3)    NULL,
    ISpAliq     DECIMAL(8,4)  NULL,
    vBCIS       DECIMAL(15,2) NULL,
    vIS         DECIMAL(15,2) NULL
GO

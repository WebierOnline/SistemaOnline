-- Rodar no SQL Server Management Studio (SSMS) para atualizar a tabela NotaFiscalItens
ALTER TABLE NotaFiscalItens ADD
    -- [Bloco IBS]
    cClassTrib      VARCHAR(20)     NULL,
    IBSCBS_CST      VARCHAR(3)      NULL,
    IBS_vBC         DECIMAL(15, 2)  NOT NULL DEFAULT 0,
    IBS_UFpAliq     DECIMAL(5, 2)   NOT NULL DEFAULT 0,
    IBS_MunpAliq    DECIMAL(5, 2)   NOT NULL DEFAULT 0,
    IBS_pRed        DECIMAL(5, 2)   NOT NULL DEFAULT 0,
    IBS_vIBSUF      DECIMAL(15, 2)  NOT NULL DEFAULT 0,
    IBS_vIBSMun     DECIMAL(15, 2)  NOT NULL DEFAULT 0,
    IBS_vIBS        DECIMAL(15, 2)  NOT NULL DEFAULT 0,

    -- [Bloco CBS]
    CBS_vBC         DECIMAL(15, 2)  NOT NULL DEFAULT 0, -- Adicionado para seguran�a fiscal
    CBS_pAliq       DECIMAL(5, 2)   NOT NULL DEFAULT 0,
    CBS_pRed        DECIMAL(5, 2)   NOT NULL DEFAULT 0,
    CBS_vCBS        DECIMAL(15, 2)  NOT NULL DEFAULT 0,

    -- [Bloco IS]
    cClassTrib_IS   VARCHAR(20)     NULL,
    IS_CST          VARCHAR(2)      NULL,
    IS_tipo_calculo INT             NOT NULL DEFAULT 1, -- 1=% (Ad Valorem), 2=Fixo (Espec�fica)
    IS_vBC          DECIMAL(15, 2)  NOT NULL DEFAULT 0,
    IS_pAliq        DECIMAL(5, 2)   NOT NULL DEFAULT 0,
    IS_qUnid        DECIMAL(15, 4)  NOT NULL DEFAULT 0, -- 4 casas decimais para litragem/unidades
    IS_vUnid        DECIMAL(15, 4)  NOT NULL DEFAULT 0, -- 4 casas decimais para valor por unidade
    IS_vIS          DECIMAL(15, 2)  NOT NULL DEFAULT 0,
    uTrib_IS        VARCHAR(6)      NULL;
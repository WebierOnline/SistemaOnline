/****** 2025 ******/

IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('NotaFiscal') AND name = 'FormaPagamento'
)
    ALTER TABLE NotaFiscal ADD FormaPagamento varchar(28) NULL;
GO

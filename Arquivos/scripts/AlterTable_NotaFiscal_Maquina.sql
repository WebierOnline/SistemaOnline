IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('NotaFiscal') AND name = 'MAQUINA'
)
    ALTER TABLE NotaFiscal ADD MAQUINA NVARCHAR(20) NULL;

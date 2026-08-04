IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('NotaFiscalItens') AND name = 'Item_pedido'
)
    ALTER TABLE NotaFiscalItens ADD Item_pedido SMALLINT NULL;
GO

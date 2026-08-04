-- Adiciona Cod_Pedido em NotaFiscalItens
-- (Item já existe como sequência do item da NF)
IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('NotaFiscalItens') AND name = 'Cod_Pedido'
)
    ALTER TABLE NotaFiscalItens ADD Cod_Pedido INT NULL;
GO

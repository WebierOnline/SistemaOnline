-- Adiciona Cod_Pedido em NotaFiscalItens
-- (Item já existe como sequência do item da NF)
ALTER TABLE NotaFiscalItens ADD Cod_Pedido INT NULL;
GO

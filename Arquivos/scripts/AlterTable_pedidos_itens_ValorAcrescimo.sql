-- Rateio do acrescimo (optAscrescRS/optAscrescPorc/txtAcresc do PDV.frm, gravado em
-- pedidos.ValorAcrescReal) por item da venda, no mesmo espirito do Desconto que
-- pedidos_itens ja tem. Necessario pra alimentar TbNFCe_Itens.ValorOutras item a item
-- quando a NFCe e gerada (mesmo padrao do TbNFCe_Itens.Desconto <- pedidos_itens.Desconto).

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('pedidos_itens') AND name = 'ValorAcrescimo')
    ALTER TABLE pedidos_itens ADD ValorAcrescimo DECIMAL(15, 2) NULL DEFAULT 0;
GO

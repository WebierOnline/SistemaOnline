-- Rateio do frete (txtFrete do PDV.frm, gravado em pedidos.ValorFreteReal) por item da
-- venda, mesmo padrao de pedidos_itens.ValorAcrescimo. Necessario pra alimentar
-- TbNFCe_Itens.Valor_Frete item a item quando a NFCe e gerada (esse campo ja existe na
-- tabela e ja e lido direto pelo TransmitirNFCe/GerarItens, so nunca foi populado).

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('pedidos_itens') AND name = 'ValorFrete')
    ALTER TABLE pedidos_itens ADD ValorFrete DECIMAL(15, 2) NULL DEFAULT 0;
GO

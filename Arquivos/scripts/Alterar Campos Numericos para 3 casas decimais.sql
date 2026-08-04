ALTER TABLE pedidos_itens ALTER COLUMN PRECO decimal(16,2);
ALTER TABLE pedidos_itens ALTER COLUMN QUANTIDADE decimal(16,3);
ALTER TABLE pedidos_itens ALTER COLUMN Desconto decimal(16,2);
ALTER TABLE pedidos_itens ALTER COLUMN Subtotal decimal(16,2);
ALTER TABLE pedidos_itens ALTER COLUMN Total decimal(16,2);

ALTER TABLE pedidos ALTER COLUMN SUBTOTAL decimal(16,2);
ALTER TABLE pedidos ALTER COLUMN TOTAL decimal(16,2);
ALTER TABLE pedidos ALTER COLUMN ENTRADA decimal(16,2);
ALTER TABLE pedidos ALTER COLUMN VALOR_DESC decimal(16,2);
ALTER TABLE pedidos ALTER COLUMN VALOR_ACRESCIMO decimal(16,2);
ALTER TABLE pedidos ALTER COLUMN ValorDescReal decimal(16,2);
ALTER TABLE pedidos ALTER COLUMN ValorAcrescReal decimal(16,2);

ALTER TABLE produtos ALTER COLUMN quant_estoque decimal(16,3);
ALTER TABLE produtos ALTER COLUMN quant_min decimal(16,3);

ALTER TABLE TbNFCe_Itens ALTER COLUMN QtdeMov decimal(16,3);
ALTER TABLE TbNFCe_Itens ALTER COLUMN desconto decimal(16,2);
ALTER TABLE TbNFCe_Itens ALTER COLUMN ValorUnit decimal(16,2);

ALTER TABLE NotaFiscalItens ALTER COLUMN ValorTotalBruto decimal(16,2);
ALTER TABLE NotaFiscalItens ALTER COLUMN ValorUnitarioComercializacao decimal(16,2);
ALTER TABLE NotaFiscalItens ALTER COLUMN vBC decimal(16,2);
ALTER TABLE NotaFiscalItens ALTER COLUMN Desconto decimal(16,2);
ALTER TABLE NotaFiscalItens ALTER COLUMN ValorDesconto decimal(16,2);
ALTER TABLE NotaFiscalItens ALTER COLUMN QuantidadeComercial decimal(16,3);

ALTER TABLE NotaFiscal ALTER COLUMN ValorProdutos decimal(16,2);
ALTER TABLE NotaFiscal ALTER COLUMN ValorNota decimal(16,2);
ALTER TABLE NotaFiscal ALTER COLUMN ValorOriginalFatura decimal(16,2);
ALTER TABLE NotaFiscal ALTER COLUMN ValorLiquidoFatura decimal(16,2);

ALTER TABLE parcelas ALTER COLUMN JUROS decimal(16,2);
ALTER TABLE parcelas ALTER COLUMN MULTA decimal(16,2);
ALTER TABLE parcelas ALTER COLUMN VALOR decimal(16,2);
ALTER TABLE parcelas ALTER COLUMN VALOR_FINAL decimal(16,2);
ALTER TABLE parcelas ALTER COLUMN DESCONTO decimal(16,2);

ALTER TABLE Produtos_Precos ALTER COLUMN CUSTO decimal(16,2);
ALTER TABLE Produtos_Precos ALTER COLUMN VALOR_VV decimal(16,2);
ALTER TABLE Produtos_Precos ALTER COLUMN VALOR_VP decimal(16,2);
ALTER TABLE Produtos_Precos ALTER COLUMN VALOR_AV decimal(16,2);
ALTER TABLE Produtos_Precos ALTER COLUMN VALOR_AP decimal(16,2);
ALTER TABLE Produtos_Precos ALTER COLUMN MARGEM_VV decimal(16,2);
ALTER TABLE Produtos_Precos ALTER COLUMN MARGEM_VP decimal(16,2);
ALTER TABLE Produtos_Precos ALTER COLUMN MARGEM_AV decimal(16,2);
ALTER TABLE Produtos_Precos ALTER COLUMN MARGEM_AP decimal(16,2);

ALTER TABLE a_receber_itens ALTER COLUMN preco decimal(16,2);

-- Notas criadas antes da atualizacao ficaram com Cod_Pedido/Item_pedido NULL em NotaFiscalItens.
-- modNFe.bas (TransmitirNFe, ~linha 449) passa
--   IIf(IsNull(NFeItens!Cod_Pedido) Or NFeItens!Cod_Pedido = 0, 0, CLng(NFeItens!Item_pedido))
--   IIf(IsNull(NFeItens!Cod_Pedido) Or NFeItens!Cod_Pedido = 0, "", CStr(NFeItens!Cod_Pedido))
-- No VB6 o IIf avalia os DOIS lados sempre, entao CLng(NULL)/CStr(NULL) dispara "Uso invalido de
-- Null" (erro 94) e a transmissao da NF-e trava, mesmo com o IsNull na condicao.
-- O exe novo ja grava 0 nessas colunas; este script so conserta as notas antigas. Idempotente.
UPDATE NotaFiscalItens SET Cod_Pedido  = 0 WHERE Cod_Pedido  IS NULL;
UPDATE NotaFiscalItens SET Item_pedido = 0 WHERE Item_pedido IS NULL;

SELECT COUNT(*) AS itens_ainda_null
FROM NotaFiscalItens WHERE Cod_Pedido IS NULL OR Item_pedido IS NULL;

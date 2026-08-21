-- Frete do PDV (novo campo txtFrete, sempre em R$, sem opcao de %). Espelha
-- ValorDescReal/ValorAcrescReal - valor real gravado no pedido pra depois alimentar
-- TbNFCe.Valor_Frete (cabecalho) via NFCeIncluir.

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('pedidos') AND name = 'ValorFreteReal')
    ALTER TABLE pedidos ADD ValorFreteReal DECIMAL(15, 2) NULL DEFAULT 0;
GO

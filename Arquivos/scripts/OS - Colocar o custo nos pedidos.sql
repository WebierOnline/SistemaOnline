IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('pedidos_itens') AND name = 'Custo')
    ALTER TABLE pedidos_itens ADD Custo decimal(16,2) NULL DEFAULT 0;
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('pedidos_itens') AND name = 'SUBTOTAL_CUSTO')
    ALTER TABLE pedidos_itens ADD SUBTOTAL_CUSTO decimal(16,2) NULL DEFAULT 0;
GO

-- Preenche Custo/SUBTOTAL_CUSTO nos itens de pedido a partir do ultimo preco cadastrado.
-- Blindado contra dado corrompido (codigo de barras gravado no campo CUSTO de Produtos_Precos):
-- o valor que sai de Produtos_Precos passa por CONVERT(float,...) (nunca estoura) e so vira
-- decimal depois de cortado na faixa valida; item fora da faixa fica com custo 0.

-- Passo 1: custo unitario
UPDATE PI
SET PI.Custo = CONVERT(decimal(12,2),
        CASE WHEN PP.v >= 0 AND PP.v < 1000000 THEN PP.v ELSE 0 END)
FROM pedidos_itens PI
CROSS APPLY (
    SELECT TOP 1 CONVERT(float, CUSTO) AS v
    FROM Produtos_Precos
    WHERE Produtos_Precos.COD_PRODUTO = PI.COD_PRODUTO
      AND Produtos_Precos.CUSTO >= 0
      AND Produtos_Precos.CUSTO < 1000000
    ORDER BY CODIGO DESC
) PP
WHERE PI.Custo = 0 OR PI.Custo IS NULL;

-- Passo 2: subtotal do custo = custo unitario * quantidade
UPDATE PI
SET PI.SUBTOTAL_CUSTO = CONVERT(decimal(16,2),
        CASE WHEN CALC.v >= 0 AND CALC.v < 10000000000000 THEN CALC.v ELSE 0 END)
FROM pedidos_itens PI
CROSS APPLY (
    SELECT CONVERT(float, ISNULL(PI.Custo,0)) * CONVERT(float, ISNULL(PI.QUANTIDADE,0)) AS v
) CALC
WHERE PI.SUBTOTAL_CUSTO IS NULL OR PI.SUBTOTAL_CUSTO = 0;

-- Verificacao imediata
SELECT COUNT(*) AS Restantes_Zerados FROM pedidos_itens WHERE Custo = 0;

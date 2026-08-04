-- Executa o update e já mostra quantos restaram logo em seguida
UPDATE PI
SET 
    PI.Custo = PP.CUSTO,
    PI.SUBTOTAL_CUSTO = (PP.CUSTO * PI.QUANTIDADE)
FROM pedidos_itens PI
CROSS APPLY (
    SELECT TOP 1 CUSTO
    FROM Produtos_Precos
    WHERE Produtos_Precos.COD_PRODUTO = PI.COD_PRODUTO
    ORDER BY CODIGO DESC
) PP
WHERE PI.Custo = 0 OR PI.Custo IS NULL;

-- Verificação imediata
SELECT COUNT(*) AS Restantes_Zerados FROM pedidos_itens WHERE Custo = 0;

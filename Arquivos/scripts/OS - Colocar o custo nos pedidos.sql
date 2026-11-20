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
--
-- Os UPDATEs rodam em LOTES de 20000 linhas (cada lote e uma transacao curta). Antes era um UPDATE unico
-- em toda a tabela e, em banco grande (300 mil+ itens), enchia o log de transacoes (Msg 9002).
-- Cada lote so pega linhas que ainda vao mudar, entao o laco termina e o script continua idempotente.
SET NOCOUNT ON
GO

-- Passo 1: custo unitario (so onde existe preco valido maior que zero)
DECLARE @n int
SET @n = 1
WHILE @n > 0
BEGIN
    UPDATE TOP (20000) PI
    SET PI.Custo = CONVERT(decimal(12,2), PP.v)
    FROM pedidos_itens PI
    CROSS APPLY (
        SELECT TOP 1 CONVERT(float, CUSTO) AS v
        FROM Produtos_Precos
        WHERE Produtos_Precos.COD_PRODUTO = PI.COD_PRODUTO
          AND Produtos_Precos.CUSTO >= 0
          AND Produtos_Precos.CUSTO < 1000000
        ORDER BY CODIGO DESC
    ) PP
    WHERE (PI.Custo = 0 OR PI.Custo IS NULL)
      AND PP.v >= 0.01 AND PP.v < 1000000;
    SET @n = @@ROWCOUNT;
    CHECKPOINT;
END
GO

-- Itens sem preco: Custo nulo vira 0
DECLARE @n int
SET @n = 1
WHILE @n > 0
BEGIN
    UPDATE TOP (20000) pedidos_itens SET Custo = 0 WHERE Custo IS NULL;
    SET @n = @@ROWCOUNT;
    CHECKPOINT;
END
GO

-- Passo 2: subtotal do custo = custo unitario * quantidade
DECLARE @n int
SET @n = 1
WHILE @n > 0
BEGIN
    UPDATE TOP (20000) PI
    SET PI.SUBTOTAL_CUSTO = CONVERT(decimal(16,2), CALC.v)
    FROM pedidos_itens PI
    CROSS APPLY (
        SELECT CONVERT(float, ISNULL(PI.Custo,0)) * CONVERT(float, ISNULL(PI.QUANTIDADE,0)) AS v
    ) CALC
    WHERE (PI.SUBTOTAL_CUSTO IS NULL OR PI.SUBTOTAL_CUSTO = 0)
      AND CALC.v >= 0.01 AND CALC.v < 10000000000000;
    SET @n = @@ROWCOUNT;
    CHECKPOINT;
END
GO

-- Subtotal nulo vira 0
DECLARE @n int
SET @n = 1
WHILE @n > 0
BEGIN
    UPDATE TOP (20000) pedidos_itens SET SUBTOTAL_CUSTO = 0 WHERE SUBTOTAL_CUSTO IS NULL;
    SET @n = @@ROWCOUNT;
    CHECKPOINT;
END
GO

SET NOCOUNT OFF
GO

-- Verificacao imediata
SELECT COUNT(*) AS Restantes_Zerados FROM pedidos_itens WHERE Custo = 0;

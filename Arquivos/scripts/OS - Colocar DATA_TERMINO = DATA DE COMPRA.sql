-- Preenche OS.DATA_TERMINO com a data de compra do pedido ligado. Guardado contra tabela inexistente.
IF OBJECT_ID('OS', 'U') IS NOT NULL AND OBJECT_ID('pedidos', 'U') IS NOT NULL
    UPDATE os
    SET os.DATA_TERMINO = p.DATA_COMPRA
    FROM OS os
    INNER JOIN pedidos p ON os.COD_PEDIDO = p.COD_PEDIDO;
GO

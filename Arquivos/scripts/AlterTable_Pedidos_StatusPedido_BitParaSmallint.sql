-- pedidos.status_pedido esta como BIT (so aceita 0/1), mas o codigo sempre tratou como inteiro:
--   0 = aberto, 1 = fechado, -1 = pausado (PDV.frm cmdAvanVendaPausar_Click / Estonar.frm cboStatus "PAUSADO"), 3 = referenciado em Estonar (filtro "incompleto")
-- Como BIT so guarda 0/1, gravar -1 e silenciosamente convertido para 1 pelo SQL Server (qualquer valor <> 0 vira 1).
-- Resultado: pausar uma venda marcava o pedido como "fechado" (status=1) em vez de "pausado", e toda busca
-- por status_pedido = -1 (Reiniciar Venda no PDV, filtro PAUSADO no Estonar) nunca encontrava nada.
IF EXISTS (
    SELECT 1 FROM sys.columns c
    JOIN sys.types t ON c.system_type_id = t.system_type_id AND c.user_type_id = t.user_type_id
    WHERE c.object_id = OBJECT_ID('pedidos') AND c.name = 'status_pedido' AND t.name = 'bit'
)
BEGIN
    ALTER TABLE pedidos ALTER COLUMN status_pedido SMALLINT NULL;
END

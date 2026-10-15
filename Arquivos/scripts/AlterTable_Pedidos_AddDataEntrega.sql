-- PDV.frm (cmdFinalizar_Click) grava pedidos.DATA_ENTREGA ao finalizar a venda, mas essa
-- coluna nunca foi migrada pra clientes antigos (so existia no banco-modelo). Sem ela, o
-- PDV trava com "Nome de coluna 'DATA_ENTREGA' invalido" ao tentar finalizar qualquer venda.
IF COL_LENGTH('pedidos', 'DATA_ENTREGA') IS NULL
    ALTER TABLE pedidos ADD DATA_ENTREGA DATETIME NULL;
GO

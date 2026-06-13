data = open(r'C:\Projeto\OnlineCommerce\Forms\NFe_Completa.frm', 'rb').read()
text = data.decode('windows-1252')

# ── Patch 1: simplificar CASE WHEN Status — remover referência a NotaFiscal.cod_pedido
old1 = ("(CASE WHEN NotaFiscal.cod_pedido = pedidos.cod_pedido OR EXISTS"
        "(SELECT 1 FROM NotaFiscalItens WHERE Cod_Pedido = pedidos.cod_pedido AND Cod_Pedido > 0)"
        " THEN 'SIM' ELSE 'N\xc3O' END) AS Status")
new1 = ("(CASE WHEN EXISTS"
        "(SELECT 1 FROM NotaFiscalItens WHERE Cod_Pedido = pedidos.cod_pedido AND Cod_Pedido > 0)"
        " THEN 'SIM' ELSE 'N\xc3O' END) AS Status")

# ── Patch 2: remover LEFT JOIN NotaFiscal do FROM de cada query
old2 = ('"FROM cliente INNER JOIN pedidos ON cliente.codigo = pedidos.cod_cliente'
        ' LEFT JOIN NotaFiscal ON NotaFiscal.cod_pedido = pedidos.cod_pedido " & _')
new2 = ('"FROM cliente INNER JOIN pedidos ON cliente.codigo = pedidos.cod_cliente " & _')

# ── Case Else: FROM e WHERE na mesma linha (sem " & _")
old2b = ('"FROM cliente INNER JOIN pedidos ON cliente.codigo = pedidos.cod_cliente'
         ' LEFT JOIN NotaFiscal ON NotaFiscal.cod_pedido = pedidos.cod_pedido WHERE pedidos.cod_pedido = \'0\';"')
new2b = ('"FROM cliente INNER JOIN pedidos ON cliente.codigo = pedidos.cod_cliente'
         ' WHERE pedidos.cod_pedido = \'0\';"')

c1  = text.count(old1)
c2  = text.count(old2)
c2b = text.count(old2b)
print(f'Patch1 (CASE WHEN): {c1}  |  Patch2a (LEFT JOIN): {c2}  |  Patch2b (Case Else): {c2b}')

if c1 == 7 and c2 == 6 and c2b == 1:
    text = text.replace(old1, new1)
    text = text.replace(old2, new2)
    text = text.replace(old2b, new2b)
    out = text.encode('windows-1252')
    out = out.replace(b'\r\n', b'\n').replace(b'\r', b'\n').replace(b'\n', b'\r\n')
    open(r'C:\Projeto\OnlineCommerce\Forms\NFe_Completa.frm', 'wb').write(out)
    print('OK — 7+6+1 substituições')
else:
    print('ABORTADO — contagens inesperadas')

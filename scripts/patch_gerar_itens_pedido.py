data = open(r'C:\Projeto\Compartilhado\Modulos\modNFe.bas', 'rb').read()
text = data.decode('windows-1252')

# ── Patch: GerarItens NFe — parâmetros 23 (nItemPedidoCompra) e 24 (numeroPedidoCompra)
# Posição atual: ..., "", "", 0, "", "", "", "", "", 1, infAdiProd, ...
#                         ^22(RECOPI) ^23(nItemPed)=0  ^24(numPed)=""  25-28=""  indTotal=1
old = (
    'NFeItens!ValorOutros, NFeItens!ValorSeguro, "", "", 0, "", "", "", "", "", 1, infAdiProd, 0, "", 0, mensagemAlerta, mensagemErro)'
)
new = (
    'NFeItens!ValorOutros, NFeItens!ValorSeguro, "", "", '
    'IIf(IsNull(NFeItens!Cod_Pedido) Or NFeItens!Cod_Pedido = 0, 0, CLng(NFeItens!Item_pedido)), '
    'IIf(IsNull(NFeItens!Cod_Pedido) Or NFeItens!Cod_Pedido = 0, "", CStr(NFeItens!Cod_Pedido)), '
    '"", "", "", "", 1, infAdiProd, 0, "", 0, mensagemAlerta, mensagemErro)'
)

c = text.count(old)
print(f'count: {c}')
if c == 1:
    text = text.replace(old, new)
    out = text.encode('windows-1252')
    out = out.replace(b'\r\n', b'\n').replace(b'\r', b'\n').replace(b'\n', b'\r\n')
    open(r'C:\Projeto\Compartilhado\Modulos\modNFe.bas', 'wb').write(out)
    print('OK')
else:
    print('ABORTADO')

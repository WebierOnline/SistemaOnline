# -*- coding: utf-8 -*-
# Fix: ROUND() explicito nos 2 UPDATEs percentuais do rateio de desconto
# (RatearDescontoPedidosItensOS/RatearDescontoServicosAutoOS), aplicados em
# rateio_os_recapadora.py. Bug achado com dados reais (Arquivos/mod64.txt,
# OS 773/pedido 1362, desconto geral 10% so em servicos, pct=22,25% exato):
# "Desconto = (Subtotal*pct/100)" e "Total = Subtotal - (Subtotal*pct/100)"
# sao 2 expressoes calculadas e arredondadas de forma INDEPENDENTE pelo SQL
# Server ao atribuir pra decimal(x,2) - quando o valor cai exatamente em
# ponto de meio-arredondamento (ex: 11.125, 33.375), cada lado pode arredondar
# "para cima" separadamente, fazendo Total != Subtotal-Desconto por 1 centavo
# em cada linha (visto nos dados: 50,00-11,13=38,87 mas Total gravou 38,88).
#
# Fix: usar ROUND(...,2) explicito, com a MESMA expressao reaproveitada nos
# 2 lados - garante Desconto+Total=Subtotal sempre, linha a linha.
path = 'OrdemServico/Forms/OS_Recapadora.frm'
data = open(path, 'rb').read()
text = data.decode('cp1252')

def ru(text, old, new, label):
    c = text.count(old)
    if c != 1:
        raise AssertionError(label + ': esperado 1, achou ' + str(c))
    return text.replace(old, new)

old1 = 'dbData.Execute "UPDATE pedidos_itens SET Desconto = (Subtotal * " & Replace(CStr(vPct), ",", ".") & " / 100), Total = Subtotal - (Subtotal * " & Replace(CStr(vPct), ",", ".") & " / 100) WHERE (cod_pedido = " & txtCodPedido.Text & ")"'
new1 = 'dbData.Execute "UPDATE pedidos_itens SET Desconto = ROUND(Subtotal * " & Replace(CStr(vPct), ",", ".") & " / 100, 2), Total = Subtotal - ROUND(Subtotal * " & Replace(CStr(vPct), ",", ".") & " / 100, 2) WHERE (cod_pedido = " & txtCodPedido.Text & ")"'
text = ru(text, old1, new1, 'ROUND pedidos_itens')

old2 = 'dbData.Execute "UPDATE OS_Servicos_Auto SET desconto = (subtotal * " & Replace(CStr(vPct), ",", ".") & " / 100), total = subtotal - (subtotal * " & Replace(CStr(vPct), ",", ".") & " / 100) WHERE (cod_os = " & vCodOS & ")"'
new2 = 'dbData.Execute "UPDATE OS_Servicos_Auto SET desconto = ROUND(subtotal * " & Replace(CStr(vPct), ",", ".") & " / 100, 2), total = subtotal - ROUND(subtotal * " & Replace(CStr(vPct), ",", ".") & " / 100, 2) WHERE (cod_os = " & vCodOS & ")"'
text = ru(text, old2, new2, 'ROUND OS_Servicos_Auto')

data2 = text.encode('cp1252')
data2 = data2.replace(b'\r\n', b'\n').replace(b'\r', b'\n').replace(b'\n', b'\r\n')
decoded = data2.decode('cp1252')
bare_lf = decoded.replace('\r\n', '').count('\n')
print('bare_lf:', bare_lf)
open(path, 'wb').write(data2)
print('OK')

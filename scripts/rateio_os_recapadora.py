# -*- coding: utf-8 -*-
path = 'OrdemServico/Forms/OS_Recapadora.frm'
data = open(path, 'rb').read()
text = data.decode('cp1252')

def ru(text, old, new, label):
    c = text.count(old)
    if c != 1:
        raise AssertionError(label + ': esperado 1, achou ' + str(c))
    return text.replace(old, new)

nl = lambda s: s.replace('\n', '\r\n')

# 1) Novas variaveis module-level (subtotais usados no rateio)
old1 = nl('''Dim rEquip As ADODB.Recordset
''')
new1 = nl('''Dim rEquip As ADODB.Recordset

'Usados no rateio de desconto geral entre produtos/servicos no fechamento da
'OS (ver ValidarRateioDescontoOS/RatearDescontoProdutosServicosOS) - subtotal
'bruto calculado antes do BEGIN TRANSACTION e reaproveitado dentro dela
Dim vSubtotalProdOS As Currency
Dim vSubtotalServOS As Currency
''')
text = ru(text, old1, new1, '1) novas vars module-level')

# 2) Validacao do rateio, logo apos calcular varValorRealAcresc, antes do
#    bloco de transacao (BEGIN TRANSACTION)
old2 = nl('''ElseIf optAscrescPorc.Value = True Then
    If txtAcresc.Text = "0,00" Then
        varValorRealAcresc = Format(0, ocMONEY)
    Else
        varValorRealAcresc = Format(((CCur(txtSubtotal.Text) * CCur(txtAcresc.Text)) / 100), ocMONEY)
    End If
End If

'declarar quem receber os produtos''')
new2 = nl('''ElseIf optAscrescPorc.Value = True Then
    If txtAcresc.Text = "0,00" Then
        varValorRealAcresc = Format(0, ocMONEY)
    Else
        varValorRealAcresc = Format(((CCur(txtSubtotal.Text) * CCur(txtAcresc.Text)) / 100), ocMONEY)
    End If
End If

'validar o rateio do desconto geral entre produtos/servicos (frmDescAplicado:
'optTudo/optProdutos/optServicos) - so se aplica quando o desconto e geral
'(txtDesc habilitado); quando ja tem desconto individual por item/servico,
'nao ha nada pra validar (ValidarRateioDescontoOS sai sem fazer nada)
If Not ValidarRateioDescontoOS(CCur(varValorRealDesc)) Then
    cmdFinalizar.Enabled = True
    Exit Sub
End If

'declarar quem receber os produtos''')
text = ru(text, old2, new2, '2) validacao antes da transacao')

# 3) Chamada do rateio - caminho A PRAZO (logo depois do UPDATE pedidos)
SP11 = ' ' * 11
old3 = nl('''                 "status_pedido = 1 " & _
                 "WHERE (cod_pedido = " & txtCodPedido.Text & ");"
              dbData.Execute sSQL
''' + SP11 + '''
           'COM ENTRADA =========================================================================''')
new3 = nl('''                 "status_pedido = 1 " & _
                 "WHERE (cod_pedido = " & txtCodPedido.Text & ");"
              dbData.Execute sSQL
''' + SP11 + '''
           'rateia o desconto geral entre produtos/servicos (pula se for
           'desconto individual) - ver ValidarRateioDescontoOS acima
           RatearDescontoProdutosServicosOS CCur(varValorRealDesc)
''' + SP11 + '''
           'COM ENTRADA =========================================================================''')
text = ru(text, old3, new3, '3) chamada rateio A PRAZO')

# 4) Chamada do rateio - caminho A VISTA (logo depois do UPDATE pedidos)
old4 = nl('''              "status_pedido = 1 " & _
              "WHERE (cod_pedido = " & txtCodPedido.Text & ");"
           dbData.Execute sSQL
''' + SP11 + '''
           '===========================================CRIAR E DAR BAIXA EM PARCELAS ==========================================''')
new4 = nl('''              "status_pedido = 1 " & _
              "WHERE (cod_pedido = " & txtCodPedido.Text & ");"
           dbData.Execute sSQL
''' + SP11 + '''
           'rateia o desconto geral entre produtos/servicos (pula se for
           'desconto individual) - ver ValidarRateioDescontoOS acima
           RatearDescontoProdutosServicosOS CCur(varValorRealDesc)
''' + SP11 + '''
           '===========================================CRIAR E DAR BAIXA EM PARCELAS ==========================================''')
text = ru(text, old4, new4, '4) chamada rateio A VISTA')

# 5) Novas Subs/Functions, inseridas antes de cmdFinalizar_Click
old5 = nl('''Private Sub cmdFinalizar_Click()''')
new5 = nl('''\'============================================================================
\'RATEIO DE DESCONTO ENTRE PRODUTOS E SERVICOS (fechamento da OS)
\'Adaptado do padrao RatearDescontoItensPedido do PDV (PDV.frm). So se aplica
\'quando o desconto informado no fechamento (txtDesc/txtDescDinheiro) e GERAL,
\'ou seja, txtDesc.Enabled = True. Quando ja existe desconto individual dado
\'por item/servico (txtDescGeral > 0), o sistema trava o txtDesc espelhando a
\'soma (ver os pontos que fazem txtDesc.Enabled = False) - nesse caso o
\'rateio NAO roda, pra nao estragar a precisao que ja esta correta.
\'
\'frmDescAplicado (optTudo/optProdutos/optServicos) decide como distribuir:
\' - optTudo:      rateia proporcional entre produtos e servicos (pelo
\'                 subtotal bruto de cada lado)
\' - optProdutos:  desconto 100% nos produtos (servicos fica sem desconto)
\' - optServicos:  desconto 100% nos servicos (produtos fica sem desconto)
\'============================================================================

\'Soma o subtotal bruto dos produtos (pedidos_itens) e servicos
\'(OS_Servicos_Auto) do pedido/OS atual - usado pra validar e pra calcular o
\'percentual do rateio. Le ANTES do BEGIN TRANSACTION (nunca dentro).
Private Sub CarregarSubtotaisRateioOS()
Dim rSub As ADODB.Recordset

Set rSub = dbData.OpenRecordset("SELECT ISNULL(SUM(Subtotal),0) AS vSoma FROM pedidos_itens WHERE (cod_pedido = " & txtCodPedido.Text & ")")
vSubtotalProdOS = ValidateNull(rSub("vSoma"))
rSub.Close

Set rSub = dbData.OpenRecordset("SELECT ISNULL(SUM(subtotal),0) AS vSoma FROM OS_Servicos_Auto WHERE (cod_os = " & vCodOS & ")")
vSubtotalServOS = ValidateNull(rSub("vSoma"))
rSub.Close
Set rSub = Nothing
End Sub

\'Confere se o modo de rateio escolhido e compativel com o valor do desconto -
\'evita desconto maior que o subtotal da categoria escolhida (daria
\'desconto/total negativo no item). So valida quando o desconto e geral
\'(txtDesc habilitado); quando e individual, nao ha rateio, nao ha o que
\'validar.
Private Function ValidarRateioDescontoOS(ByVal vValorDesc As Currency) As Boolean
ValidarRateioDescontoOS = True
If txtDesc.Enabled = False Then Exit Function

CarregarSubtotaisRateioOS

If optProdutos.Value = True Then
    If vValorDesc > vSubtotalProdOS Then
        ShowMsg "O desconto (" & Format(vValorDesc, ocMONEY) & ") \xe9 maior que o subtotal dos PRODUTOS (" & Format(vSubtotalProdOS, ocMONEY) & "). Marque ""Tudo"" no desconto aplicado, ou ajuste o valor do desconto.", vbExclamation
        ValidarRateioDescontoOS = False
    End If
ElseIf optServicos.Value = True Then
    If vValorDesc > vSubtotalServOS Then
        ShowMsg "O desconto (" & Format(vValorDesc, ocMONEY) & ") \xe9 maior que o subtotal dos SERVI\xc7OS (" & Format(vSubtotalServOS, ocMONEY) & "). Marque ""Tudo"" no desconto aplicado, ou ajuste o valor do desconto.", vbExclamation
        ValidarRateioDescontoOS = False
    End If
End If
End Function

\'Aplica o desconto geral (vValorDesc) nos itens/servicos, conforme
\'frmDescAplicado. Chamado DENTRO da transacao do fechamento, logo depois do
\'UPDATE da tabela pedidos. Pula completamente quando o desconto e
\'individual (txtDesc.Enabled = False).
Private Sub RatearDescontoProdutosServicosOS(ByVal vValorDesc As Currency)
If txtDesc.Enabled = False Then Exit Sub

Dim vValorAlvoProd As Currency, vValorAlvoServ As Currency
Dim vPctProd As Double, vPctServ As Double
Dim vSubtotalTudo As Currency

If optTudo.Value = True Then
    vSubtotalTudo = vSubtotalProdOS + vSubtotalServOS
    If vSubtotalTudo <> 0 And vValorDesc <> 0 Then
        vValorAlvoProd = FormatNumber(vValorDesc * (vSubtotalProdOS / vSubtotalTudo), 2)
        vValorAlvoServ = vValorDesc - vValorAlvoProd
    Else
        vValorAlvoProd = 0
        vValorAlvoServ = 0
    End If
ElseIf optProdutos.Value = True Then
    vValorAlvoProd = vValorDesc
    vValorAlvoServ = 0
ElseIf optServicos.Value = True Then
    vValorAlvoProd = 0
    vValorAlvoServ = vValorDesc
End If

If vSubtotalProdOS <> 0 Then vPctProd = (vValorAlvoProd / vSubtotalProdOS) * 100 Else vPctProd = 0
If vSubtotalServOS <> 0 Then vPctServ = (vValorAlvoServ / vSubtotalServOS) * 100 Else vPctServ = 0

RatearDescontoPedidosItensOS vPctProd, vValorAlvoProd
RatearDescontoServicosAutoOS vPctServ, vValorAlvoServ
End Sub

\'Rateia o percentual vPct em cada pedidos_itens do pedido (proporcional ao
\'subtotal de cada item), e ajusta o item de maior CODIGO pra bater
\'exatamente com vValorAlvo - mesmo padrao do RatearDescontoItensPedido do
\'PDV (PDV.frm).
Private Sub RatearDescontoPedidosItensOS(ByVal vPct As Double, ByVal vValorAlvo As Currency)
Dim vSomaDescItens As Currency
Dim rSub As ADODB.Recordset

dbData.Execute "UPDATE pedidos_itens SET Desconto = (Subtotal * " & Replace(CStr(vPct), ",", ".") & " / 100), Total = Subtotal - (Subtotal * " & Replace(CStr(vPct), ",", ".") & " / 100) WHERE (cod_pedido = " & txtCodPedido.Text & ")"

If vPct <> 0 Then
    Set rSub = dbData.OpenRecordset("SELECT ISNULL(SUM(Desconto),0) AS vSoma FROM pedidos_itens WHERE (cod_pedido = " & txtCodPedido.Text & ")")
    vSomaDescItens = FormatNumber(ValidateNull(rSub("vSoma")), 2)
    rSub.Close
    Set rSub = Nothing

    If vValorAlvo <> vSomaDescItens Then
        dbData.Execute "UPDATE pedidos_itens SET Desconto = Desconto + (" & Replace(CStr(vValorAlvo - vSomaDescItens), ",", ".") & "), Total = Total - (" & Replace(CStr(vValorAlvo - vSomaDescItens), ",", ".") & ") WHERE (CODIGO = (SELECT MAX(CODIGO) FROM pedidos_itens WHERE cod_pedido = " & txtCodPedido.Text & "))"
    End If
End If
End Sub

\'Mesma logica de RatearDescontoPedidosItensOS, mas pra OS_Servicos_Auto
\'(colunas em minusculo: desconto/total/subtotal).
Private Sub RatearDescontoServicosAutoOS(ByVal vPct As Double, ByVal vValorAlvo As Currency)
Dim vSomaDescServ As Currency
Dim rSub As ADODB.Recordset

dbData.Execute "UPDATE OS_Servicos_Auto SET desconto = (subtotal * " & Replace(CStr(vPct), ",", ".") & " / 100), total = subtotal - (subtotal * " & Replace(CStr(vPct), ",", ".") & " / 100) WHERE (cod_os = " & vCodOS & ")"

If vPct <> 0 Then
    Set rSub = dbData.OpenRecordset("SELECT ISNULL(SUM(desconto),0) AS vSoma FROM OS_Servicos_Auto WHERE (cod_os = " & vCodOS & ")")
    vSomaDescServ = FormatNumber(ValidateNull(rSub("vSoma")), 2)
    rSub.Close
    Set rSub = Nothing

    If vValorAlvo <> vSomaDescServ Then
        dbData.Execute "UPDATE OS_Servicos_Auto SET desconto = desconto + (" & Replace(CStr(vValorAlvo - vSomaDescServ), ",", ".") & "), total = total - (" & Replace(CStr(vValorAlvo - vSomaDescServ), ",", ".") & ") WHERE (codigo = (SELECT MAX(codigo) FROM OS_Servicos_Auto WHERE cod_os = " & vCodOS & "))"
    End If
End If
End Sub

Private Sub cmdFinalizar_Click()''')
text = ru(text, old5, new5, '5) novas Subs/Function')

data2 = text.encode('cp1252')
data2 = data2.replace(b'\r\n', b'\n').replace(b'\r', b'\n').replace(b'\n', b'\r\n')
decoded = data2.decode('cp1252')
bare_lf = decoded.replace('\r\n', '').count('\n')
print('bare_lf:', bare_lf)
open(path, 'wb').write(data2)
print('OK, bytes:', len(data2))

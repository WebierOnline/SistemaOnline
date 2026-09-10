# -*- coding: utf-8 -*-
"""PDV.frm - correcoes 1..5 do cashback (analise 2026-09-09):
 1) mascara de data yyyy-dd-MM -> yyyy-MM-dd (DATA_ABATIDO e VALIDADE)
 2) pCashbackPercentual (String) -> Double locale-safe antes da conta
 4) vCashbackValidade fica Date puro (sem roundtrip String->Date por locale)
 3) F7 nao deixa o cashback-desconto passar do subtotal da venda (clamp + aviso)
 5) lstCashBack nao fica grudado num cliente antigo (esconde ao trocar cliente / nova venda / load)
.frm cp1252 -> edicao binaria + CRLF.
"""
import sys

p = r"C:\projeto\PDV\Forms\PDV.frm"
d = open(p, "rb").read().replace(b"\r\n", b"\n")

A = b"\xe1"  # a-agudo cp1252
E = b"\xe9"  # e-agudo cp1252

repls = [
    # 1) DATA_ABATIDO no UPDATE de baixa
    (b'DATA_ABATIDO = \'" & Format$(Date, "yyyy-dd-MM") & "\'',
     b'DATA_ABATIDO = \'" & Format$(Date, "yyyy-MM-dd") & "\''),

    # 2 + 4) parse do percentual + validade como Date puro
    (b"        'valor do cashback\n"
     b"        vValorVenda = CCur(txtTotalDesc.Text)\n"
     b"        ValorCash = (vValorVenda * pCashbackPercentual) / 100\n"
     b"        \n"
     b"        'validade do cashback\n"
     b'        vCashbackValidade = Format(DateAdd("d", Val(vCashbackLimite), Date), "dd/mm/yy")\n',
     b"        'valor do cashback\n"
     b"        Dim dPercCash As Double\n"
     b'        dPercCash = Val(Replace(Trim(pCashbackPercentual), ",", "."))   \'config vem como texto ("1,5")\n'
     b"        vValorVenda = CCur(txtTotalDesc.Text)\n"
     b"        ValorCash = (vValorVenda * dPercCash) / 100\n"
     b"        \n"
     b"        'validade do cashback (Date puro - sem roundtrip String->Date por locale)\n"
     b'        vCashbackValidade = DateAdd("d", Val(vCashbackLimite), Date)\n'),

    # 1) VALIDADE no INSERT
    (b'Format$(vCashbackValidade, "yyyy-dd-MM")',
     b'Format$(vCashbackValidade, "yyyy-MM-dd")'),

    # 3) F7: clamp do cashback ao subtotal da venda
    (b"        If Not r Is Nothing Then\n"
     b"            optDescRS.Value = True\n"
     b'            txtDesc.Text = FormatNumber(ValidateNull(r("vValorSomaCash")), 2)\n'
     b"            If txtDesc.Text > 0 Then vUsandoCashBack = True Else vUsandoCashBack = False\n",
     b"        If Not r Is Nothing Then\n"
     b"            Dim dCashSoma As Double, dCashSub As Double\n"
     b'            If IsNull(r("vValorSomaCash")) Then dCashSoma = 0 Else dCashSoma = CDbl(r("vValorSomaCash"))\n'
     b'            dCashSub = Val(Replace(Replace(txtSubtotal.Text, ".", ""), ",", "."))\n'
     b"            If dCashSub > 0 And dCashSoma > dCashSub Then\n"
     b'                MsgBox "Saldo de cashback (" & FormatNumber(dCashSoma, 2) & ") maior que o valor da venda." & vbCrLf & _\n'
     b'                       "Ser' + A + b' abatido at' + E + b' o total da venda (" & FormatNumber(dCashSub, 2) & "); o saldo restante ser' + A + b' consumido.", vbInformation, "Cashback"\n'
     b"                dCashSoma = dCashSub\n"
     b"            End If\n"
     b"            optDescRS.Value = True\n"
     b"            txtDesc.Text = FormatNumber(dCashSoma, 2)\n"
     b"            If dCashSoma > 0 Then vUsandoCashBack = True Else vUsandoCashBack = False\n"),

    # 5a) troca de cliente
    (b"Private Sub CboCliente_LostFocus()\n"
     b"On Error GoTo TrataErro\n",
     b"Private Sub CboCliente_LostFocus()\n"
     b"On Error GoTo TrataErro\n"
     b"\n"
     b"lstCashBack.Visible = False   'cliente pode ter mudado - nao deixar o painel de cashback grudado no anterior\n"),

    # 5b) apos finalizar venda
    (b'    vTipoEdicao = ""\n'
     b"End If\n"
     b"\n"
     b"vUsandoCashBack = False\n"
     b"Exit Sub\n",
     b'    vTipoEdicao = ""\n'
     b"End If\n"
     b"\n"
     b"vUsandoCashBack = False\n"
     b"lstCashBack.Visible = False\n"
     b"Exit Sub\n"),

    # 5c) Form_Load / reset geral
    (b"PesoF4 = False\n"
     b"vUsandoCashBack = False\n",
     b"PesoF4 = False\n"
     b"vUsandoCashBack = False\n"
     b"lstCashBack.Visible = False\n"),

    # 5d) cancelar pedido
    (b"        If txtCodBarra.Enabled = True Then txtCodBarra.SetFocus\n"
     b"    End If\n"
     b"End If\n"
     b"vUsandoCashBack = False\n"
     b"End Sub\n",
     b"        If txtCodBarra.Enabled = True Then txtCodBarra.SetFocus\n"
     b"    End If\n"
     b"End If\n"
     b"vUsandoCashBack = False\n"
     b"lstCashBack.Visible = False\n"
     b"End Sub\n"),
]

for i, (old, new) in enumerate(repls, 1):
    if new in d and old not in d:
        print(f"#{i} ja aplicado"); continue
    n = d.count(old)
    if n != 1:
        print(f"#{i}: {n} ocorrencias (esperado 1) -- ABORTA"); sys.exit(1)
    d = d.replace(old, new, 1)
    print(f"#{i} OK")

open(p, "wb").write(d.replace(b"\n", b"\r\n"))
print("gravado")

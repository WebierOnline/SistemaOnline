# -*- coding: utf-8 -*-
"""PDV.frm - selecao de cashbacks no frmCashBack (checkboxes):
 - lstCashBack agora vive dentro do Frame frmCashBack -> F6 e os resets alternam frmCashBack.Visible
 - coluna CODIGO oculta (Width 0); lstCashBack.Checkboxes = True
 - cmdMarcarTodos: marca/desmarca todas + troca o Caption
 - cmdUsarEscolhidos: soma so os marcados no desconto e guarda os CODIGOs em vCashbackCodigosUsados
 - AplicarCashbackVenda: baixa so nos CODIGOs marcados (vazio = todos, compat. com F7)
.frm cp1252 -> edicao binaria + CRLF.
"""
import sys

p = r"C:\projeto\PDV\Forms\PDV.frm"
d = open(p, "rb").read().replace(b"\r\n", b"\n")
A = b"\xe1"   # a agudo
E = b"\xe9"   # e agudo
O = b"\xd3"   # O agudo maiusculo (usado em "CODIGO"/"COD.VENDA" dos headers)

reps = []

# --- var de modulo -------------------------------------------------------------
reps.append((
b"Public vCashbackLimite As String        'cashback Limite\n",
b"Public vCashbackLimite As String        'cashback Limite\n"
b'Public vCashbackCodigosUsados As String  \'CODIGOs de Pedidos_Cashback marcados no frmCashBack (vazio = usar todos)\n',
))

# --- F6a: condicao do toggle -------------------------------------------------
reps.append((
b'        If lstCashBack.Visible = False Then\n'
b'            dbData.Execute "UPDATE Pedidos_Cashback SET INVALIDO = 1 WHERE (COD_CLIENTE = " & txtCodCliente.Text & ") and ABATIDO = 0 and INVALIDO = 0 and VALIDADE < CONVERT(DATETIME, CONVERT(date, GETDATE()));"',
b'        If frmCashBack.Visible = False Then\n'
b'            dbData.Execute "UPDATE Pedidos_Cashback SET INVALIDO = 1 WHERE (COD_CLIENTE = " & txtCodCliente.Text & ") and ABATIDO = 0 and INVALIDO = 0 and VALIDADE < CONVERT(DATETIME, CONVERT(date, GETDATE()));"',
))

# --- F6b: mostra o frame + liga checkboxes ---------------------------------
reps.append((
b'            lstCashBack.Visible = True\n'
b'            Dim ListaCash As ListItem\n'
b'            lstCashBack.FullRowSelect = True\n'
b'            lstCashBack.LabelEdit = lvwManual\n'
b'            lstCashBack.Visible = True\n'
b'            lstCashBack.View = lvwReport\n'
b'            lstCashBack.HideSelection = False\n',
b'            frmCashBack.Visible = True\n'
b'            frmCashBack.ZOrder 0\n'
b'            lstCashBack.Visible = True\n'
b'            Dim ListaCash As ListItem\n'
b'            lstCashBack.Checkboxes = True\n'
b'            lstCashBack.FullRowSelect = True\n'
b'            lstCashBack.LabelEdit = lvwManual\n'
b'            lstCashBack.View = lvwReport\n'
b'            lstCashBack.HideSelection = False\n',
))

# --- F6c: oculta coluna CODIGO + reseta caption do botao ------------------
reps.append((
b'            lstCashBack.ColumnHeaders.Add , , "VALIDADE", 1200\n'
b'            \n'
b'            sSQL = "SELECT CODIGO, COD_PEDIDO, VALOR_VENDA, VALOR_CASHBACK, VALIDADE " & _\n',
b'            lstCashBack.ColumnHeaders.Add , , "VALIDADE", 1200\n'
b"            lstCashBack.ColumnHeaders(1).Width = 0   'oculta a coluna CODIGO (usada so internamente)\n"
b'            cmdMarcarTodos.Caption = "Marcar Todos"\n'
b'            \n'
b'            sSQL = "SELECT CODIGO, COD_PEDIDO, VALOR_VENDA, VALOR_CASHBACK, VALIDADE " & _\n',
))

# --- F6d: ramo Else (esconder) -------------------------------------------
reps.append((
b'        Else\n'
b'            lstCashBack.Visible = False\n'
b'        End If\n'
b'    End If\n'
b'    End If\n'
b'ElseIf KeyCode = vbKeyF7 Then',
b'        Else\n'
b'            frmCashBack.Visible = False\n'
b'        End If\n'
b'    End If\n'
b'    End If\n'
b'ElseIf KeyCode = vbKeyF7 Then',
))

# --- F7: limpa a selecao (F7 = usar todo o saldo) ------------------------
reps.append((
b'ElseIf KeyCode = vbKeyF7 Then\n'
b'    If txtCodCliente <> "1" And txtCodCliente.Text <> "" Then\n'
b'    If vCashbackAV = "SIM" Or vCashbackAP = "SIM" Then\n'
b'        dbData.Execute "UPDATE Pedidos_Cashback SET INVALIDO = 1 WHERE (COD_CLIENTE = " & txtCodCliente.Text & ") and ABATIDO = 0 and INVALIDO = 0 and VALIDADE < CONVERT(DATETIME, CONVERT(date, GETDATE()));"',
b'ElseIf KeyCode = vbKeyF7 Then\n'
b'    vCashbackCodigosUsados = ""   \'F7 = usar todo o saldo disponivel\n'
b'    If txtCodCliente <> "1" And txtCodCliente.Text <> "" Then\n'
b'    If vCashbackAV = "SIM" Or vCashbackAP = "SIM" Then\n'
b'        dbData.Execute "UPDATE Pedidos_Cashback SET INVALIDO = 1 WHERE (COD_CLIENTE = " & txtCodCliente.Text & ") and ABATIDO = 0 and INVALIDO = 0 and VALIDADE < CONVERT(DATETIME, CONVERT(date, GETDATE()));"',
))

# --- resets do ciclo de venda (troca cliente / cancelar / finalizar / load) --
reps.append((
b"lstCashBack.Visible = False   'cliente pode ter mudado - nao deixar o painel de cashback grudado no anterior\n",
b"frmCashBack.Visible = False   'cliente pode ter mudado - nao deixar o painel de cashback grudado no anterior\n"
b'vCashbackCodigosUsados = ""\n',
))
reps.append((
b"vUsandoCashBack = False\nlstCashBack.Visible = False\nEnd Sub\n",
b'vUsandoCashBack = False\nfrmCashBack.Visible = False\nvCashbackCodigosUsados = ""\nEnd Sub\n',
))
reps.append((
b"vUsandoCashBack = False\nlstCashBack.Visible = False\nExit Sub\n",
b'vUsandoCashBack = False\nfrmCashBack.Visible = False\nvCashbackCodigosUsados = ""\nExit Sub\n',
))
reps.append((
b"PesoF4 = False\nvUsandoCashBack = False\nlstCashBack.Visible = False\n",
b'PesoF4 = False\nvUsandoCashBack = False\nfrmCashBack.Visible = False\nvCashbackCodigosUsados = ""\n',
))

# --- AplicarCashbackVenda: baixa so nos marcados -------------------------
reps.append((
b'        If vUsandoCashBack = True Then\n'
b'            dbData.Execute "UPDATE Pedidos_Cashback SET VALOR_ABATIDO = VALOR_CASHBACK, ABATIDO = 1, DATA_ABATIDO = \'" & Format$(Date, "yyyymmdd") & "\', COD_PEDIDOABATIDO = " & txtCodPedido.Text & ", COD_FUNCIONARIO = " & txtCodFuncAP.Text & " WHERE (COD_CLIENTE = " & txtCodCliente.Text & ") and ABATIDO = 0 and INVALIDO = 0;"\n'
b'        End If\n',
b'        If vUsandoCashBack = True Then\n'
b'            Dim sFiltroCodCash As String\n'
b'            If Trim(vCashbackCodigosUsados) <> "" Then sFiltroCodCash = " and CODIGO in (" & vCashbackCodigosUsados & ")" Else sFiltroCodCash = ""\n'
b'            dbData.Execute "UPDATE Pedidos_Cashback SET VALOR_ABATIDO = VALOR_CASHBACK, ABATIDO = 1, DATA_ABATIDO = \'" & Format$(Date, "yyyymmdd") & "\', COD_PEDIDOABATIDO = " & txtCodPedido.Text & ", COD_FUNCIONARIO = " & txtCodFuncAP.Text & " WHERE (COD_CLIENTE = " & txtCodCliente.Text & ") and ABATIDO = 0 and INVALIDO = 0" & sFiltroCodCash & ";"\n'
b'        End If\n',
))

# --- novos handlers (depois de AplicarCashbackVenda) -------------------
novos = (
b'End Sub\n'
b'\n'
b'Private Sub cmdUsarEscolhidos_Click()\n'
b"   'soma so os cashbacks marcados no lstCashBack e guarda os CODIGOs pra baixa seletiva\n"
b'   Dim dSomaCB As Double, dSubCB As Double\n'
b'   Dim sCodsCB As String\n'
b'   dSomaCB = 0\n'
b'   sCodsCB = ""\n'
b'   For i = 1 To lstCashBack.ListItems.Count\n'
b'      If lstCashBack.ListItems(i).Checked Then\n'
b'         dSomaCB = dSomaCB + Val(Replace(Replace(lstCashBack.ListItems(i).ListSubItems(3).Text, ".", ""), ",", "."))\n'
b'         sCodsCB = sCodsCB & IIf(sCodsCB = "", "", ",") & lstCashBack.ListItems(i).Text\n'
b'      End If\n'
b'   Next i\n'
b'\n'
b'   If sCodsCB = "" Then\n'
b'      MsgBox "Marque ao menos um cashback para usar.", vbExclamation, "Cashback"\n'
b'      Exit Sub\n'
b'   End If\n'
b'\n'
b'   dSubCB = Val(Replace(Replace(txtSubtotal.Text, ".", ""), ",", "."))\n'
b'   If dSubCB > 0 And dSomaCB > dSubCB Then\n'
b'      MsgBox "Cashback marcado (" & FormatNumber(dSomaCB, 2) & ") maior que o valor da venda." & vbCrLf & _\n'
b'             "Ser' + A + b' abatido at' + E + b' o total da venda (" & FormatNumber(dSubCB, 2) & "); o saldo restante dos marcados ser' + A + b' consumido.", vbInformation, "Cashback"\n'
b'      dSomaCB = dSubCB\n'
b'   End If\n'
b'\n'
b'   optDescRS.Value = True\n'
b'   txtDesc.Text = FormatNumber(dSomaCB, 2)\n'
b'   vUsandoCashBack = (dSomaCB > 0)\n'
b'   vCashbackCodigosUsados = sCodsCB\n'
b'   frmCashBack.Visible = False\n'
b'End Sub\n'
b'\n'
b'Private Sub cmdMarcarTodos_Click()\n'
b'   Dim bTodas As Boolean\n'
b'   If lstCashBack.ListItems.Count = 0 Then Exit Sub\n'
b'   bTodas = True\n'
b'   For i = 1 To lstCashBack.ListItems.Count\n'
b'      If Not lstCashBack.ListItems(i).Checked Then bTodas = False: Exit For\n'
b'   Next i\n'
b'   For i = 1 To lstCashBack.ListItems.Count\n'
b'      lstCashBack.ListItems(i).Checked = Not bTodas\n'
b'   Next i\n'
b'   If bTodas Then cmdMarcarTodos.Caption = "Marcar Todos" Else cmdMarcarTodos.Caption = "Desmarcar Todos"\n'
b'End Sub\n'
b'\n'
b'Private Sub lstCashBack_ItemCheck(ByVal Item As MSComctlLib.ListItem)\n'
b"   'mantem o caption do cmdMarcarTodos coerente quando o usuario marca/desmarca na mao\n"
b'   Dim bTodas As Boolean\n'
b'   bTodas = (lstCashBack.ListItems.Count > 0)\n'
b'   For i = 1 To lstCashBack.ListItems.Count\n'
b'      If Not lstCashBack.ListItems(i).Checked Then bTodas = False: Exit For\n'
b'   Next i\n'
b'   If bTodas Then cmdMarcarTodos.Caption = "Desmarcar Todos" Else cmdMarcarTodos.Caption = "Marcar Todos"\n'
b'End Sub\n'
b'\n'
b'Private Function NFCeJaExisteParaPedido(ByVal pCodPedido As String) As Boolean\n'
)
reps.append((
b'End Sub\n\nPrivate Function NFCeJaExisteParaPedido(ByVal pCodPedido As String) As Boolean\n',
novos,
))

for i, (old, new) in enumerate(reps, 1):
    if new in d and old not in d:
        print(f"#{i} ja aplicado"); continue
    n = d.count(old)
    if n != 1:
        print(f"#{i}: {n} ocorrencias (esperado 1) -- ABORTA")
        sys.exit(1)
    d = d.replace(old, new, 1)
    print(f"#{i} OK")

open(p, "wb").write(d.replace(b"\n", b"\r\n"))
print("gravado")

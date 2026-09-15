# -*- coding: utf-8 -*-
"""Estonar.frm - Mostrar_Pedido, item 6 da ordem de ataque: janela de data default
quando cboCriterios = "NENHUM" (antes trazia o historico inteiro, cada linha com as
subqueries -> travava em loja com anos de venda).

- Mostrar_Pedido, ramo NENHUM: se mskData tem data valida ->
  `and pedidos.DATA_COMPRA >= CONVERT(DATETIME, 'yyyymmdd')` (sem separador = ISO,
  imune a DATEFORMAT do server do cliente - ver feedback_sqlserver_data_sem_separador).
  Campo vazio/invalido -> sem limite (opt-in explicito do usuario pra "tudo").
- cboCriterios_LostFocus, ramo NENHUM: passa a MOSTRAR lblData/mskData/cmdCal1 com
  caption "A partir de:" e default = hoje - 90 dias (so preenche se estiver vazio).
- cboCriterios_LostFocus, ramo DATA: reseta lblData.Caption = "Data:" (a NENHUM mexe nele).

Disparo de Mostrar_Pedido = so cmdExibir_Click e cmdExcluirPedido_Click (nao em change de
combo), entao o usuario ve o campo "A partir de:" preenchido antes de clicar Exibir.

.frm cp1252 -> edicao binaria + normaliza CRLF.
"""
import sys

p = r"C:\projeto\PDV\Forms\Estonar.frm"
d = open(p, "rb").read().replace(b"\r\n", b"\n").replace(b"\r", b"\n")

reps = []

# ---------------------------------------------------- 1) Mostrar_Pedido ramo NENHUM
o1 = (b'If cboCriterios.Text = "NENHUM" Then\n'
      b'    varCriterio = " "\n'
      b'ElseIf cboCriterios.Text = "C\xd3D. PEDIDO" Then\n')
n1 = (b'If cboCriterios.Text = "NENHUM" Then\n'
      b'    If IsDate(mskData.Text) Then\n'
      b'        varCriterio = " and pedidos.DATA_COMPRA >= CONVERT(DATETIME, \'" & Format$(CDate(mskData.Text), "yyyymmdd") & "\') "\n'
      b'    Else\n'
      b'        varCriterio = " "\n'
      b'    End If\n'
      b'ElseIf cboCriterios.Text = "C\xd3D. PEDIDO" Then\n')
reps.append(("Mostrar_Pedido / NENHUM -> janela de data", o1, n1))

# ---------------------------------------------------- 2) cboCriterios_LostFocus ramo DATA
o2 = (b'    lblData.Visible = True\n'
      b'    mskData.Visible = True\n'
      b'    cmdCal1.Visible = True\n'
      b'    lblProduto.Visible = False\n')
n2 = (b'    lblData.Caption = "Data:"\n'
      b'    lblData.Visible = True\n'
      b'    mskData.Visible = True\n'
      b'    cmdCal1.Visible = True\n'
      b'    lblProduto.Visible = False\n')
reps.append(("cboCriterios_LostFocus / DATA -> reseta caption", o2, n2))

# ---------------------------------------------------- 3) cboCriterios_LostFocus ramo NENHUM
o3 = (b'    lblData.Visible = False\n'
      b'    mskData.Visible = False\n'
      b'    cmdCal1.Visible = False\n'
      b'    lblProduto.Visible = False\n'
      b'    cboProduto.Visible = False\n'
      b'    lblCodBarra.Visible = False\n'
      b'    txtCodBarra.Visible = False\n'
      b'ElseIf cboCriterios.Text = "C\xd3D. PEDIDO" Then\n')
n3 = (b'    lblData.Caption = "A partir de:"\n'
      b'    lblData.Visible = True\n'
      b'    mskData.Visible = True\n'
      b'    cmdCal1.Visible = True\n'
      b'    If Not IsDate(mskData.Text) Then mskData.Text = Format$(DateAdd("d", -90, Date), "dd/mm/yyyy")\n'
      b'    lblProduto.Visible = False\n'
      b'    cboProduto.Visible = False\n'
      b'    lblCodBarra.Visible = False\n'
      b'    txtCodBarra.Visible = False\n'
      b'ElseIf cboCriterios.Text = "C\xd3D. PEDIDO" Then\n')
reps.append(("cboCriterios_LostFocus / NENHUM -> mostra campo 'A partir de:'", o3, n3))

for nome, old, new in reps:
    if new in d and old not in d:
        print("[ja] " + nome); continue
    c = d.count(old)
    if c != 1:
        print("[ABORTA] %s -- %d ocorrencias de old" % (nome, c)); sys.exit(1)
    d = d.replace(old, new, 1)
    print("[ok]  " + nome)

open(p, "wb").write(d.replace(b"\n", b"\r\n"))
print("gravado " + p)

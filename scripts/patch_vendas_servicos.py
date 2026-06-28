# -*- coding: utf-8 -*-
"""Patch: adiciona bloco POR SERVICOS completo em Vendas_Consulta_PorProdutos.frm"""
import sys

FILE = r'C:\Projeto\OnlineCommerce\Forms\Vendas_Consulta_PorProdutos.frm'

with open(FILE, 'rb') as f:
    raw = f.read()

raw = raw.replace(b'\r\n', b'\n').replace(b'\r', b'\n')
text = raw.decode('windows-1252')

changes = []

# ------------------------------------------------------------------
# 1. Adiciona INDICE para POR SERVICOS no cmdLocalizar_Click
# ------------------------------------------------------------------
old1 = (
    "'INDICE\n"
    "If cboTipo.Text = \"POR PRODUTOS\" Then\n"
    "   If cboIndice.Text = \"QUANT.\" Then\n"
    "      INDICE = \"quantidade ;\"\n"
    "   ElseIf cboIndice.Text = \"PRODUTO\" Then\n"
    "      INDICE = \"produtos.descricao ;\"\n"
    "   ElseIf cboIndice.Text = \"DATA\" Then\n"
    "      INDICE = \"pedidos_itens.data ;\"\n"
    "   ElseIf cboIndice.Text = \"PEDIDO\" Then\n"
    "      INDICE = \"pedidos_itens.cod_pedido ;\"\n"
    "   Else\n"
    "      INDICE = \"produtos.descricao ;\"\n"
    "   End If\n"
    "End If\n"
)
new1 = old1 + (
    "If cboTipo.Text = \"POR SERVIÇOS\" Then\n"
    "   If cboIndice.Text = \"QUANT.\" Then\n"
    "      INDICE = \"s.quantidade ;\"\n"
    "   ElseIf cboIndice.Text = \"PRODUTO\" Then\n"
    "      INDICE = \"s.descricao ;\"\n"
    "   ElseIf cboIndice.Text = \"DATA\" Then\n"
    "      INDICE = \"s.data ;\"\n"
    "   ElseIf cboIndice.Text = \"PEDIDO\" Then\n"
    "      INDICE = \"s.cod_os ;\"\n"
    "   Else\n"
    "      INDICE = \"s.descricao ;\"\n"
    "   End If\n"
    "End If\n"
)
changes.append((old1, new1, '1 - INDICE POR SERVICOS'))

# ------------------------------------------------------------------
# 2. Adiciona ESPECIFICO/MENSAL e ESPECIFICO em cboCriterioPrinc_LostFocus
# ------------------------------------------------------------------
old2 = (
    "ElseIf cboCriterioPrinc.Text = \"DATA\" Then\n"
    "    lblInicio.Visible = True\n"
    "    lblInicio.Caption = \"Data\"\n"
    "    mskInicio.Visible = True\n"
    "    lblFim.Visible = False\n"
    "    mskFim.Visible = False\n"
    "    lblAte.Visible = False\n"
    "    cmdCalendario1.Visible = True\n"
    "    cmdCalendario2.Visible = False\n"
    "    lblMes.Visible = False\n"
    "    cboMes.Visible = False\n"
    "    lblAno.Visible = False\n"
    "    cboAno.Visible = False\n"
    "End If\n"
)
new2 = (
    "ElseIf cboCriterioPrinc.Text = \"DATA\" Then\n"
    "    lblInicio.Visible = True\n"
    "    lblInicio.Caption = \"Data\"\n"
    "    mskInicio.Visible = True\n"
    "    lblFim.Visible = False\n"
    "    mskFim.Visible = False\n"
    "    lblAte.Visible = False\n"
    "    cmdCalendario1.Visible = True\n"
    "    cmdCalendario2.Visible = False\n"
    "    lblMes.Visible = False\n"
    "    cboMes.Visible = False\n"
    "    lblAno.Visible = False\n"
    "    cboAno.Visible = False\n"
    "ElseIf cboCriterioPrinc.Text = \"ESPECIFICO/MENSAL\" Then\n"
    "    lblInicio.Visible = False\n"
    "    mskInicio.Visible = False\n"
    "    lblFim.Visible = False\n"
    "    mskFim.Visible = False\n"
    "    lblAte.Visible = False\n"
    "    cmdCalendario1.Visible = False\n"
    "    cmdCalendario2.Visible = False\n"
    "    lblMes.Visible = True\n"
    "    cboMes.Visible = True\n"
    "    lblAno.Visible = True\n"
    "    cboAno.Visible = True\n"
    "ElseIf cboCriterioPrinc.Text = \"ESPECIFICO\" Then\n"
    "    lblInicio.Visible = False\n"
    "    mskInicio.Visible = False\n"
    "    lblFim.Visible = False\n"
    "    mskFim.Visible = False\n"
    "    lblAte.Visible = False\n"
    "    cmdCalendario1.Visible = False\n"
    "    cmdCalendario2.Visible = False\n"
    "    lblMes.Visible = False\n"
    "    cboMes.Visible = False\n"
    "    lblAno.Visible = False\n"
    "    cboAno.Visible = False\n"
    "End If\n"
)
changes.append((old2, new2, '2 - ESPECIFICO/MENSAL e ESPECIFICO em cboCriterioPrinc_LostFocus'))

# ------------------------------------------------------------------
# 3. cboCriterioSec_GotFocus — diferencia POR SERVICOS
# ------------------------------------------------------------------
old3 = (
    "Private Sub cboCriterioSec_GotFocus()\n"
    "cboCriterioSec.Clear\n"
    "\n"
    "cboCriterioSec.AddItem \"DESCRIÇÃO\"\n"
    "cboCriterioSec.AddItem \"CÓD. BARRA\"\n"
    "cboCriterioSec.AddItem \"REFERÊNCIA\"\n"
    "cboCriterioSec.AddItem \"FABRICANTE\"\n"
    "cboCriterioSec.AddItem \"CATEGORIA\"\n"
    "\n"
    "moCombo.AttachTo cboCriterioSec\n"
    "End Sub\n"
)
new3 = (
    "Private Sub cboCriterioSec_GotFocus()\n"
    "cboCriterioSec.Clear\n"
    "\n"
    "If cboTipo.Text = \"POR SERVIÇOS\" Then\n"
    "   cboCriterioSec.AddItem \"DESCRIÇÃO\"\n"
    "Else\n"
    "   cboCriterioSec.AddItem \"DESCRIÇÃO\"\n"
    "   cboCriterioSec.AddItem \"CÓD. BARRA\"\n"
    "   cboCriterioSec.AddItem \"REFERÊNCIA\"\n"
    "   cboCriterioSec.AddItem \"FABRICANTE\"\n"
    "   cboCriterioSec.AddItem \"CATEGORIA\"\n"
    "End If\n"
    "\n"
    "moCombo.AttachTo cboCriterioSec\n"
    "End Sub\n"
)
changes.append((old3, new3, '3 - cboCriterioSec_GotFocus POR SERVICOS'))

# ------------------------------------------------------------------
# 4. cboDescricao_GotFocus — carrega servicos quando POR SERVICOS
# ------------------------------------------------------------------
old4 = (
    "Private Sub cboDescricao_GotFocus()\n"
    "   Dim sSQL As String\n"
    "   Dim r As ADODB.Recordset\n"
    "   \n"
    "   cboDescricao.Clear\n"
    "   \n"
    "If cboCriterioSec.Text = \"DESCRIÇÃO\" Then\n"
)
new4 = (
    "Private Sub cboDescricao_GotFocus()\n"
    "   Dim sSQL As String\n"
    "   Dim r As ADODB.Recordset\n"
    "   \n"
    "   cboDescricao.Clear\n"
    "   \n"
    "If cboTipo.Text = \"POR SERVIÇOS\" Then\n"
    "   sSQL = \"SELECT DISTINCT descricao FROM OS_Servicos_Auto ORDER BY descricao;\"\n"
    "   Set r = dbData.OpenRecordset(sSQL)\n"
    "   Do While Not r.EOF\n"
    "      cboDescricao.AddItem r(\"descricao\")\n"
    "      r.MoveNext\n"
    "   Loop\n"
    "   If r.State <> 0 Then r.Close\n"
    "   Set r = Nothing\n"
    "   moCombo.AttachTo cboDescricao\n"
    "   Exit Sub\n"
    "End If\n"
    "\n"
    "If cboCriterioSec.Text = \"DESCRIÇÃO\" Then\n"
)
changes.append((old4, new4, '4 - cboDescricao_GotFocus POR SERVICOS'))

# ------------------------------------------------------------------
# 5. Substitui bloco POR SERVICOS incompleto em cmdLocalizar_Click
# ------------------------------------------------------------------
old5 = (
    "ElseIf cboTipo.Text = \"POR SERVIÇOS\" Then\n"
    "         'TODOS\n"
    "         If cboCriterioPrinc.Text = \"TODOS\" And cboCriterioSec.Text = \"\" Then\n"
    "            sSQL = \"SELECT os_servicos.cod_produto, os_servicos.descricao as var_desc, SUM(os_servicos.quantidade) AS var_qtde, preco, SUM(preco * quantidade) AS var_total \" & _\n"
    "               \"FROM produtos LEFT JOIN os_servicos ON produtos.codigo = os_servicos.cod_produto \" & _\n"
    "               \"LEFT JOIN pedidos ON os_servicos.cod_pedido = pedidos.cod_pedido \" & _\n"
    "               \"WHERE (pedidos.tipo_pedido = 'BALCAO' or pedidos.tipo_pedido = 'OFICINA')  \" & _\n"
    "               \"GROUP BY os_servicos.cod_produto, produtos.descricao, produtos.tamanho, produtos.fabricante, produtos.ref, os_servicos.preco ORDER BY \" & INDICE\n"
    "         End If\n"
    "End If\n"
)
new5 = (
    "ElseIf cboTipo.Text = \"POR SERVIÇOS\" Then\n"
    "   Dim sBase As String\n"
    "   sBase = \"SELECT s.codigo, OS.COD_OS AS varCodPed, s.data AS varData, s.descricao AS varNome, \" & _\n"
    "           \"s.preco AS varValor, s.quantidade AS varQuant, s.subtotal AS varSubtotal, \" & _\n"
    "           \"s.desconto AS varDesc, s.total AS varTotal, 0 AS var_CodOS \" & _\n"
    "           \"FROM OS_Servicos_Auto s INNER JOIN OS ON s.cod_os = OS.COD_OS \"\n"
    "\n"
    "   If cboCriterioPrinc.Text = \"TODOS\" Then\n"
    "      If cboCriterioSec.Text = \"DESCRIÇÃO\" Then\n"
    "         If cboDescricao.Text = \"\" Then Exit Sub\n"
    "         sSQL = sBase & \"WHERE s.descricao = '\" & cboDescricao.Text & \"' ORDER BY \" & INDICE\n"
    "      Else\n"
    "         sSQL = sBase & \"ORDER BY \" & INDICE\n"
    "      End If\n"
    "   ElseIf cboCriterioPrinc.Text = \"MENSAL\" Then\n"
    "      If cboMes.Text = \"\" Or cboAno.Text = \"\" Then Exit Sub\n"
    "      If cboCriterioSec.Text = \"DESCRIÇÃO\" Then\n"
    "         If cboDescricao.Text = \"\" Then Exit Sub\n"
    "         sSQL = sBase & \"WHERE s.descricao = '\" & cboDescricao.Text & \"' AND MONTH(s.data) = \" & cboMes.ListIndex + 1 & \" AND YEAR(s.data) = \" & cboAno & \" ORDER BY \" & INDICE\n"
    "      Else\n"
    "         sSQL = sBase & \"WHERE MONTH(s.data) = \" & cboMes.ListIndex + 1 & \" AND YEAR(s.data) = \" & cboAno & \" ORDER BY \" & INDICE\n"
    "      End If\n"
    "   ElseIf cboCriterioPrinc.Text = \"ESPECIFICO/MENSAL\" Then\n"
    "      If cboDescricao.Text = \"\" Then Exit Sub\n"
    "      If cboMes.Text = \"\" Or cboAno.Text = \"\" Then Exit Sub\n"
    "      sSQL = sBase & \"WHERE s.descricao = '\" & cboDescricao.Text & \"' AND MONTH(s.data) = \" & cboMes.ListIndex + 1 & \" AND YEAR(s.data) = \" & cboAno & \" ORDER BY \" & INDICE\n"
    "   ElseIf cboCriterioPrinc.Text = \"ESPECIFICO\" Then\n"
    "      If cboDescricao.Text = \"\" Then Exit Sub\n"
    "      sSQL = sBase & \"WHERE s.descricao = '\" & cboDescricao.Text & \"' ORDER BY \" & INDICE\n"
    "   End If\n"
    "End If\n"
)
changes.append((old5, new5, '5 - bloco POR SERVICOS em cmdLocalizar_Click'))

# ------------------------------------------------------------------
# 6. Troca chamada FormatarGrid_ProdDetalhado por condicional
# ------------------------------------------------------------------
old6 = "FormatarGrid_ProdDetalhado r\n"
new6 = (
    "If cboTipo.Text = \"POR SERVIÇOS\" Then\n"
    "   FormatarGrid_Servicos r\n"
    "Else\n"
    "   FormatarGrid_ProdDetalhado r\n"
    "End If\n"
)
changes.append((old6, new6, '6 - FormatarGrid condicional'))

# ------------------------------------------------------------------
# 7. Insere novo sub FormatarGrid_Servicos antes de FormatarGrid_ProdDetalhado
# ------------------------------------------------------------------
new_sub = (
    "Private Sub FormatarGrid_Servicos(rTabela As ADODB.Recordset)\n"
    "   Dim i As Integer\n"
    "picAguarde.Visible = True\n"
    "DoEvents\n"
    "   With Grid\n"
    "      .Clear\n"
    "      .Cols = 11\n"
    "      .rows = 2\n"
    "      \n"
    "      .ColWidth(0) = 0\n"
    "      .ColWidth(1) = 750\n"
    "      .ColWidth(2) = 900\n"
    "      .ColWidth(3) = 0\n"
    "      .ColWidth(4) = 5400\n"
    "      .ColWidth(5) = 800\n"
    "      .ColWidth(6) = 800\n"
    "      .ColWidth(7) = 800\n"
    "      .ColWidth(8) = 700\n"
    "      .ColWidth(9) = 800\n"
    "      .ColWidth(10) = 0\n"
    "      \n"
    "      .TextMatrix(0, 1) = \"OS\"\n"
    "      .TextMatrix(0, 2) = \"DATA\"\n"
    "      .TextMatrix(0, 3) = \"\"\n"
    "      .TextMatrix(0, 4) = \"DESCRIÇÃO\"\n"
    "      .TextMatrix(0, 5) = \"VALOR\"\n"
    "      .TextMatrix(0, 6) = \"QTDE\"\n"
    "      .TextMatrix(0, 7) = \"=\"\n"
    "      .TextMatrix(0, 8) = \"DESC.\"\n"
    "      .TextMatrix(0, 9) = \"TOTAL\"\n"
    "      .TextMatrix(0, 10) = \"\"\n"
    "      \n"
    "      .Redraw = False\n"
    "      \n"
    "      For i = 0 To .Cols - 1\n"
    "         .Col = i\n"
    "         .Row = 0\n"
    "         .CellFontBold = True\n"
    "      Next\n"
    "      \n"
    "      .ColAlignment(1) = 1\n"
    "      \n"
    "      For i = 0 To .Cols - 1\n"
    "         .Row = 0\n"
    "         .Col = i\n"
    "         .CellAlignment = flexAlignCenterCenter\n"
    "      Next\n"
    "      \n"
    "      If Not rTabela Is Nothing Then\n"
    "         Do While Not rTabela.EOF\n"
    "            .TextMatrix(.rows - 1, 1) = Format(rTabela(\"varCodPed\"), \"000000\")\n"
    "            .TextMatrix(.rows - 1, 2) = Format(rTabela(\"varData\"), \"dd/mm/yy\")\n"
    "            .TextMatrix(.rows - 1, 3) = \"\"\n"
    "            .TextMatrix(.rows - 1, 4) = rTabela(\"varNome\")\n"
    "            .TextMatrix(.rows - 1, 5) = Format(rTabela(\"varValor\"), ocMONEY)\n"
    "            .TextMatrix(.rows - 1, 6) = rTabela(\"varQuant\")\n"
    "            .TextMatrix(.rows - 1, 7) = Format(rTabela(\"varSubtotal\"), ocMONEY)\n"
    "            .TextMatrix(.rows - 1, 8) = Format(rTabela(\"varDesc\"), ocMONEY)\n"
    "            .TextMatrix(.rows - 1, 9) = Format(rTabela(\"varTotal\"), ocMONEY)\n"
    "            .TextMatrix(.rows - 1, 10) = \"0\"\n"
    "            \n"
    "            rTabela.MoveNext\n"
    "            .rows = .rows + 1\n"
    "         Loop\n"
    "      End If\n"
    "      \n"
    "      .rows = .rows - 1\n"
    "      .Redraw = True\n"
    "   End With\n"
    "   \n"
    "   lblQtda.Caption = SomaGrid(Grid, 6)\n"
    "   lblTotal.Caption = Format(SomaGrid(Grid, 9), ocMONEY)\n"
    "picAguarde.Visible = False\n"
    "End Sub\n"
    "\n"
)
old7 = "Private Sub FormatarGrid_ProdDetalhado("
new7 = new_sub + old7
changes.append((old7, new7, '7 - novo sub FormatarGrid_Servicos'))

# ------------------------------------------------------------------
# Aplicar e verificar
# ------------------------------------------------------------------
for old, new, label in changes:
    count = text.count(old)
    if count != 1:
        print(f'ERRO [{label}]: encontrado {count} ocorrencias (esperado 1)')
        sys.exit(1)
    text = text.replace(old, new)
    print(f'OK: {label}')

# Re-encode com CRLF
text = text.replace('\r\n', '\n').replace('\r', '\n')
out = text.encode('windows-1252')
out = out.replace(b'\n', b'\r\n')

with open(FILE, 'wb') as f:
    f.write(out)

print('\nArquivo gravado com sucesso.')

import sys

data = open(r'C:\Projeto\OnlineCommerce\Forms\NFe_Completa.frm', 'rb').read()
text = data.decode('windows-1252')

# --- Patch 1a: ColWidth(0) visivel ---
old1a = '      .ColWidth(0) = 0\r\n       .ColWidth(1) = 900\r\n       .ColWidth(2) = 900\r\n       .ColWidth(3) = 1100\r\n       .ColWidth(4) = 4000\r\n       .ColWidth(5) = 2000\r\n       .ColWidth(6) = 2000'
new1a = '      .ColWidth(0) = 420\r\n       .ColWidth(1) = 900\r\n       .ColWidth(2) = 900\r\n       .ColWidth(3) = 1100\r\n       .ColWidth(4) = 4000\r\n       .ColWidth(5) = 2000\r\n       .ColWidth(6) = 2000'

# --- Patch 1b: inicializar imagem desmarcada em cada linha ---
old1b = (
    '            .TextMatrix(.rows - 1, 1) = rTabela("var_codped")\r\n'
    '            .TextMatrix(.rows - 1, 2) = rTabela("status")\r\n'
    '            .TextMatrix(.rows - 1, 3) = Format(rTabela("var_dtcompra"), "dd/mm/yy")\r\n'
    '            .TextMatrix(.rows - 1, 4) = rTabela("var_Nome")\r\n'
    '            .TextMatrix(.rows - 1, 5) = ValidateNull(rTabela("var_tipoPGTO"))\r\n'
    '            .TextMatrix(.rows - 1, 6) = Format(rTabela("var_total"), ocMONEY)\r\n'
    '            \r\n'
    '            rTabela.MoveNext\r\n'
    '            .rows = .rows + 1\r\n'
    '            i = i + 1'
)
new1b = (
    '            .TextMatrix(.rows - 1, 1) = rTabela("var_codped")\r\n'
    '            .TextMatrix(.rows - 1, 2) = rTabela("status")\r\n'
    '            .TextMatrix(.rows - 1, 3) = Format(rTabela("var_dtcompra"), "dd/mm/yy")\r\n'
    '            .TextMatrix(.rows - 1, 4) = rTabela("var_Nome")\r\n'
    '            .TextMatrix(.rows - 1, 5) = ValidateNull(rTabela("var_tipoPGTO"))\r\n'
    '            .TextMatrix(.rows - 1, 6) = Format(rTabela("var_total"), ocMONEY)\r\n'
    '            .Row = .rows - 1: .Col = 0\r\n'
    '            Set .CellPicture = imgDesmarcada.Picture\r\n'
    '            .CellPictureAlignment = 4\r\n'
    '            \r\n'
    '            rTabela.MoveNext\r\n'
    '            .rows = .rows + 1\r\n'
    '            i = i + 1'
)

# --- Patch 2: cmdConverterNFe com loop multi-selecao ---
old2 = (
    'Private Sub cmdConverterNFe_Click()\r\n'
    'If vTipoEdicaoNFe = "Novo" Or vTipoEdicaoNFe = "Edicao" Then MsgBox "Existem um NFe em aberto, Salve-a ou Cancele-a!", vbExclamation, "Online Commerce": Frm_NF.Tab = 0: Exit Sub\r\n'
    'If GridPedidos.Row = 0 Then MsgBox "Selecione uma nota fiscal na lista!", vbInformation, "Aviso do Sistema": Exit Sub\r\n'
    'If GridPedidos.TextMatrix(GridPedidos.Row, 2) = "SIM" Then MsgBox "Esse pedido j\xe1 foi transformado em NFe!", vbInformation, "Online Commerce": Exit Sub\r\n'
    'If ShowMsg("Deseja realmente transformar o pedido: " & GridPedidos.TextMatrix(GridPedidos.Row, 1) & " em Nota Fiscal?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then Exit Sub\r\n'
    '\r\n'
    'txtCodPedido.Text = (GridPedidos.TextMatrix(GridPedidos.Row, 1))\r\n'
    '\r\n'
    'vTipoEdicaoNFe = "Edicao"\r\n'
    'picAguarde2.Visible = True\r\n'
    'GravarPedido\r\n'
    'AtualizarTotaisNota\r\n'
    'txtInfComple.Text = "NFe referente a venda N\xba " & txtCodPedido.Text\r\n'
    'Frm_NF.Tab = 0\r\n'
    'cmdNovo.Enabled = False\r\n'
    'cmdSalvar.Enabled = True\r\n'
    'cmdCancelar.Enabled = True\r\n'
    'frmNota.Enabled = True\r\n'
    'frmDestinatario.Enabled = True\r\n'
    'frmItens.Enabled = True\r\n'
    'Tab_Totais.Enabled = True\r\n'
    'Tab_Produtos.Enabled = True\r\n'
    'cmdRecalcular_Click\r\n'
    'cmdExibirPedidos_Click\r\n'
    'picAguarde2.Visible = False\r\n'
    'End Sub'
)
new2 = (
    'Private Sub cmdConverterNFe_Click()\r\n'
    'If vTipoEdicaoNFe = "Novo" Or vTipoEdicaoNFe = "Edicao" Then MsgBox "Existe uma NFe em aberto, salve ou cancele antes!", vbExclamation, "Online Commerce": Frm_NF.Tab = 0: Exit Sub\r\n'
    '\r\n'
    'Dim nMarcados As Integer, i As Integer\r\n'
    'For i = 1 To GridPedidos.Rows - 1\r\n'
    '    If GridPedidos.TextMatrix(i, 0) = "1" And GridPedidos.TextMatrix(i, 2) <> "SIM" Then\r\n'
    '        nMarcados = nMarcados + 1\r\n'
    '    End If\r\n'
    'Next i\r\n'
    '\r\n'
    'If nMarcados = 0 Then MsgBox "Marque ao menos um pedido ainda n\xe3o convertido.", vbInformation, "Online Commerce": Exit Sub\r\n'
    'If ShowMsg("Deseja converter " & nMarcados & " pedido(s) em Nota Fiscal?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then Exit Sub\r\n'
    '\r\n'
    'picAguarde2.Visible = True\r\n'
    '\r\n'
    'For i = 1 To GridPedidos.Rows - 1\r\n'
    '    If GridPedidos.TextMatrix(i, 0) = "1" And GridPedidos.TextMatrix(i, 2) <> "SIM" Then\r\n'
    '        txtCodPedido.Text = GridPedidos.TextMatrix(i, 1)\r\n'
    '        vTipoEdicaoNFe = "Edicao"\r\n'
    '        GravarPedido\r\n'
    '        AtualizarTotaisNota\r\n'
    '        txtInfComple.Text = "NFe referente a venda N\xba " & txtCodPedido.Text\r\n'
    '    End If\r\n'
    'Next i\r\n'
    '\r\n'
    'vTipoEdicaoNFe = ""\r\n'
    'cmdNovo.Enabled = True\r\n'
    'cmdSalvar.Enabled = False\r\n'
    'cmdCancelar.Enabled = False\r\n'
    'frmNota.Enabled = False\r\n'
    'frmDestinatario.Enabled = False\r\n'
    'frmItens.Enabled = False\r\n'
    'Tab_Totais.Enabled = False\r\n'
    'Tab_Produtos.Enabled = False\r\n'
    'cmdExibirPedidos_Click\r\n'
    'picAguarde2.Visible = False\r\n'
    'End Sub'
)

# --- Patch 3: adicionar GridPedidos_Click antes de FormatarGridPedidos ---
old3 = 'Private Sub FormatarGridPedidos(rTabela As ADODB.Recordset)'
new3 = (
    'Private Sub GridPedidos_Click()\r\n'
    'Dim iRow As Integer\r\n'
    'iRow = GridPedidos.Row\r\n'
    'If iRow < 1 Then Exit Sub\r\n'
    'If GridPedidos.Col <> 0 Then Exit Sub\r\n'
    'If GridPedidos.TextMatrix(iRow, 0) = "1" Then\r\n'
    '    GridPedidos.TextMatrix(iRow, 0) = ""\r\n'
    '    GridPedidos.Row = iRow: GridPedidos.Col = 0\r\n'
    '    Set GridPedidos.CellPicture = imgDesmarcada.Picture\r\n'
    '    GridPedidos.CellPictureAlignment = 4\r\n'
    'Else\r\n'
    '    GridPedidos.TextMatrix(iRow, 0) = "1"\r\n'
    '    GridPedidos.Row = iRow: GridPedidos.Col = 0\r\n'
    '    Set GridPedidos.CellPicture = ImgMarcada.Picture\r\n'
    '    GridPedidos.CellPictureAlignment = 4\r\n'
    'End If\r\n'
    'End Sub\r\n'
    '\r\n'
    'Private Sub FormatarGridPedidos(rTabela As ADODB.Recordset)'
)

c1a = text.count(old1a)
c1b = text.count(old1b)
c2  = text.count(old2)
c3  = text.count(old3)
print(f'1a:{c1a}  1b:{c1b}  2:{c2}  3:{c3}')

if c1a == 1 and c1b == 1 and c2 == 1 and c3 == 1:
    text = text.replace(old1a, new1a)
    text = text.replace(old1b, new1b)
    text = text.replace(old2,  new2)
    text = text.replace(old3,  new3)
    out = text.encode('windows-1252')
    out = out.replace(b'\r\n', b'\n').replace(b'\r', b'\n').replace(b'\n', b'\r\n')
    open(r'C:\Projeto\OnlineCommerce\Forms\NFe_Completa.frm', 'wb').write(out)
    print('OK - 3 patches aplicados')
else:
    print('ERRO - algum trecho nao foi encontrado')

data = open(r'C:\Projeto\OnlineCommerce\Forms\NFe_Completa.frm', 'rb').read()
text = data.decode('windows-1252')

patches = []

# ── Patch 1: cor SIM (vermelho escuro) / NÃO (preto) na coluna 2 do GridPedidos
patches.append((
    '            .TextMatrix(.rows - 1, 2) = rTabela("status")\r\n'
    '            .TextMatrix(.rows - 1, 3) = Format(rTabela("var_dtcompra"), "dd/mm/yy")',

    '            .TextMatrix(.rows - 1, 2) = rTabela("status")\r\n'
    '            .Row = .rows - 1: .Col = 2\r\n'
    '            If rTabela("status") = "SIM" Then\r\n'
    '                .CellForeColor = RGB(139, 0, 0)\r\n'
    '            Else\r\n'
    '                .CellForeColor = vbBlack\r\n'
    '            End If\r\n'
    '            .TextMatrix(.rows - 1, 3) = Format(rTabela("var_dtcompra"), "dd/mm/yy")',
    'cor_SIM_NAO'
))

# ── Patch 2: cboClientePedidos não carrega clientes quando modo = PEDIDO
patches.append((
    'Private Sub cboClientePedidos_GotFocus()\r\n'
    'Dim r As ADODB.Recordset\r\n',

    'Private Sub cboClientePedidos_GotFocus()\r\n'
    'If cboIndicePedidos.Text = "PEDIDO" Then Exit Sub\r\n'
    'Dim r As ADODB.Recordset\r\n',
    'cbo_pedido_nao_carrega_clientes'
))

# ── Patch 3a: lblCodFabrica(9) vinculado à visibilidade das imagens (FormatarGridPedidos)
patches.append((
    '   If GridPedidos.rows > 1 Then\r\n'
    '       imgDesmarcadaTODAS.Visible = True\r\n'
    '       ImgMarcadaTODAS.Visible = False\r\n'
    '       lblCodFabrica(9).Visible = True\r\n'
    '   Else\r\n'
    '       imgDesmarcadaTODAS.Visible = False\r\n'
    '       ImgMarcadaTODAS.Visible = False\r\n'
    '       lblCodFabrica(9).Visible = False\r\n'
    '   End If',

    '   If GridPedidos.rows > 1 Then\r\n'
    '       imgDesmarcadaTODAS.Visible = True\r\n'
    '       ImgMarcadaTODAS.Visible = False\r\n'
    '   Else\r\n'
    '       imgDesmarcadaTODAS.Visible = False\r\n'
    '       ImgMarcadaTODAS.Visible = False\r\n'
    '   End If\r\n'
    '   lblCodFabrica(9).Visible = imgDesmarcadaTODAS.Visible Or ImgMarcadaTODAS.Visible',
    'lbl_visibilidade_Or'
))

# ── Patch 3b: lblCodFabrica(9) explícito no imgDesmarcadaTODAS_Click
patches.append((
    'imgDesmarcadaTODAS.Visible = False\r\n'
    'ImgMarcadaTODAS.Visible = True\r\n'
    'For i = 1 To GridPedidos.rows - 1\r\n'
    '    If GridPedidos.TextMatrix(i, 2) <> "SIM" Then\r\n'
    '        GridPedidos.RowData(i) = 1',

    'imgDesmarcadaTODAS.Visible = False\r\n'
    'ImgMarcadaTODAS.Visible = True\r\n'
    'lblCodFabrica(9).Visible = True\r\n'
    'For i = 1 To GridPedidos.rows - 1\r\n'
    '    If GridPedidos.TextMatrix(i, 2) <> "SIM" Then\r\n'
    '        GridPedidos.RowData(i) = 1',
    'lbl_desmarcadaTODAS'
))

# ── Patch 3c: lblCodFabrica(9) explícito no ImgMarcadaTODAS_Click
patches.append((
    'ImgMarcadaTODAS.Visible = False\r\n'
    'imgDesmarcadaTODAS.Visible = True\r\n'
    'For i = 1 To GridPedidos.rows - 1\r\n'
    '    If GridPedidos.TextMatrix(i, 2) <> "SIM" Then\r\n'
    '        GridPedidos.RowData(i) = 0',

    'ImgMarcadaTODAS.Visible = False\r\n'
    'imgDesmarcadaTODAS.Visible = True\r\n'
    'lblCodFabrica(9).Visible = True\r\n'
    'For i = 1 To GridPedidos.rows - 1\r\n'
    '    If GridPedidos.TextMatrix(i, 2) <> "SIM" Then\r\n'
    '        GridPedidos.RowData(i) = 0',
    'lbl_marcadaTODAS'
))

# ── Patch 4a: limpar cboMesPedidos e cboAnoPedidos no Case MENSAL
patches.append((
    '        cmdExibirPedidos.Left = 3960\r\n'
    '\r\n'
    '    Case "CLIENTE + DATAS"',

    '        cmdExibirPedidos.Left = 3960\r\n'
    '        cboMesPedidos.Text = ""\r\n'
    '        cboAnoPedidos.Text = ""\r\n'
    '\r\n'
    '    Case "CLIENTE + DATAS"',
    'limpar_MENSAL'
))

# ── Patch 4b: limpar cboMesPedidos e cboAnoPedidos no Case CLIENTE + MENSAL
patches.append((
    '        cmdExibirPedidos.Left = 10020\r\n'
    '\r\n'
    'End Select\r\n'
    'End Sub\r\n'
    '\r\n'
    'Private Sub cboIndicePedidos_GotFocus()',

    '        cmdExibirPedidos.Left = 10020\r\n'
    '        cboMesPedidos.Text = ""\r\n'
    '        cboAnoPedidos.Text = ""\r\n'
    '\r\n'
    'End Select\r\n'
    'End Sub\r\n'
    '\r\n'
    'Private Sub cboIndicePedidos_GotFocus()',
    'limpar_CLIENTE_MENSAL'
))

# ── Patch 5a: zerar Cod_Pedido ao cancelar NF
patches.append((
    '    vgDb.Execute sql\r\n'
    '    \'RsOpen TbNotas',

    '    vgDb.Execute sql\r\n'
    '    vgDb.Execute "UPDATE NotaFiscal SET Cod_Pedido = 0 WHERE CodigoNota = " & Val(vCodNota)\r\n'
    '    vgDb.Execute "UPDATE NotaFiscalItens SET Cod_Pedido = 0, Item_pedido = 0 WHERE CodigoNota = " & Val(vCodNota)\r\n'
    '    \'RsOpen TbNotas',
    'zerar_cancelar'
))

# ── Patch 5b: zerar Cod_Pedido ao inutilizar NF
patches.append((
    '      sSQL = "UPDATE NotaFiscal SET Enviada = 1, Inutilizada = 1 WHERE CodigoNota = " & IdNFProd\r\n'
    '      vgDb.Execute sSQL\r\n'
    '      MsgBox CStr(cStat)',

    '      sSQL = "UPDATE NotaFiscal SET Enviada = 1, Inutilizada = 1 WHERE CodigoNota = " & IdNFProd\r\n'
    '      vgDb.Execute sSQL\r\n'
    '      vgDb.Execute "UPDATE NotaFiscal SET Cod_Pedido = 0 WHERE CodigoNota = " & IdNFProd\r\n'
    '      vgDb.Execute "UPDATE NotaFiscalItens SET Cod_Pedido = 0, Item_pedido = 0 WHERE CodigoNota = " & IdNFProd\r\n'
    '      MsgBox CStr(cStat)',
    'zerar_inutilizar'
))

# ── Verificar contagens ──────────────────────────────────────────────────────
all_ok = True
for old, new, name in patches:
    c = text.count(old)
    status = 'OK' if c == 1 else f'ERRO (count={c})'
    print(f'{name}: {status}')
    if c != 1:
        all_ok = False

if all_ok:
    for old, new, _ in patches:
        text = text.replace(old, new)
    out = text.encode('windows-1252')
    out = out.replace(b'\r\n', b'\n').replace(b'\r', b'\n').replace(b'\n', b'\r\n')
    open(r'C:\Projeto\OnlineCommerce\Forms\NFe_Completa.frm', 'wb').write(out)
    print('\nOK — todos os patches aplicados')
else:
    print('\nABORTADO')

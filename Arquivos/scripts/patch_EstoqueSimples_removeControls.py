path = r'C:\Projeto\OnlineCommerce\Forms\Produtos_Estoque_Simples.frm'
with open(path, 'rb') as f:
    data = f.read()
content = data.decode('windows-1252')

errors = []

def sub(label, old, new, c, replace_all=False):
    n = c.count(old)
    if replace_all:
        if n == 0:
            errors.append(f'{label}: count=0')
            return c
        print(f'{label} OK ({n}x)')
        return c.replace(old, new)
    if n != 1:
        errors.append(f'{label}: count={n}')
        return c
    print(f'{label} OK')
    return c.replace(old, new, 1)

# ── 1: Remover definicoes de layout: chkDescPorIniciais, chkDescPorProduto, cboConsLinha
content = sub('layout remove 3 controls',
    '      Begin VB.CheckBox chkDescPorIniciais \r\n'
    '         Caption         =   "Por Iniciais"\r\n'
    '         ForeColor       =   &H000000FF&\r\n'
    '         Height          =   195\r\n'
    '         Left            =   1500\r\n'
    '         TabIndex        =   11\r\n'
    '         Top             =   1080\r\n'
    '         Visible         =   0   \'False\r\n'
    '         Width           =   1335\r\n'
    '      End\r\n'
    '      Begin VB.CheckBox chkDescPorProduto \r\n'
    '         Caption         =   "Por Produto"\r\n'
    '         ForeColor       =   &H000000FF&\r\n'
    '         Height          =   195\r\n'
    '         Left            =   180\r\n'
    '         TabIndex        =   10\r\n'
    '         Top             =   1080\r\n'
    '         Value           =   1  \'Checked\r\n'
    '         Visible         =   0   \'False\r\n'
    '         Width           =   1275\r\n'
    '      End\r\n'
    '      Begin VB.ComboBox cboConsLinha \r\n'
    '         Height          =   315\r\n'
    '         Left            =   120\r\n'
    '         TabIndex        =   8\r\n'
    '         Top             =   720\r\n'
    '         Visible         =   0   \'False\r\n'
    '         Width           =   3915\r\n'
    '      End\r\n',
    '',
    content)

# ── 2: Remover subs chkDescPorIniciais_Click e chkDescPorProduto_Click ────────
content = sub('remove chkDesc*_Click subs',
    'Private Sub chkDescPorIniciais_Click()\r\n'
    'If optDesc.Value = Unchecked Then Exit Sub\r\n'
    '\r\n'
    'If chkDescPorIniciais.Value = Checked Then\r\n'
    '   cboDesc.Clear\r\n'
    '   chkDescPorProduto.Value = Unchecked\r\n'
    '   cboDesc.SetFocus\r\n'
    'End If\r\n'
    'End Sub\r\n'
    '\r\n'
    'Private Sub chkDescPorProduto_Click()\r\n'
    'If optDesc.Value = Unchecked Then Exit Sub\r\n'
    '\r\n'
    'If chkDescPorProduto.Value = Checked Then\r\n'
    '   chkDescPorIniciais.Value = Unchecked\r\n'
    '   cboDesc.SetFocus\r\n'
    'End If\r\n'
    'End Sub\r\n'
    '\r\n'
    'Private Sub MostrarCriterios()',
    'Private Sub MostrarCriterios()',
    content)

# ── 3: Remover subs cboConsLinha_Click/GotFocus/LostFocus ─────────────────────
content = sub('remove cboConsLinha_* subs',
    'Private Sub cboConsLinha_Click()\r\n'
    '   \'If cboConsLinha.Text <> "" Then MostrarCriterios\r\n'
    'End Sub\r\n'
    '\r\n'
    'Private Sub cboConsLinha_GotFocus()\r\n'
    'cboConsLinha.Clear\r\n'
    '\r\n'
    'sSQL = "SELECT DISTINCT Categoria FROM Categorias WHERE Tipo_Empresa = " & tipoEmpresa & " ORDER BY Categoria;"\r\n'
    'Set r = dbData.OpenRecordset(sSQL)\r\n'
    '\r\n'
    'Do While Not r.EOF\r\n'
    '   cboConsLinha.AddItem ValidateNull(r("categoria"))\r\n'
    '   r.MoveNext\r\n'
    'Loop\r\n'
    '\r\n'
    'If r.State <> 0 Then r.Close\r\n'
    'Set r = Nothing\r\n'
    '\r\n'
    'moCombo.AttachTo cboConsLinha\r\n'
    'End Sub\r\n'
    '\r\n'
    'Private Sub cboConsLinha_LostFocus()\r\n'
    '   \'cboConsLinha_Click\r\n'
    'End Sub\r\n'
    '\r\n'
    'Private Sub Formatar_Grid_Fiscal(',
    'Private Sub Formatar_Grid_Fiscal(',
    content)

# ── 4: MostrarCriterios — substituir logica chk* por optPorPalavra/PorPalavraDupla
content = sub('MostrarCriterios optDesc logic',
    '   If chkDescPorProduto.Value = Checked Then\r\n'
    '      var_Criterio = var_Criterio & IIf(optDesc.Value, IIf(var_Criterio <> "", " AND ", "") & "produtos.descricao = \'" & cboDesc.Text & "\'", "")\r\n'
    '   ElseIf chkDescPorIniciais.Value = Checked Then\r\n'
    '      var_Criterio = Chr$(39) & cboDesc.Text & "%" & Chr(39)\r\n'
    '      var_Criterio = IIf(optDesc.Value, IIf(var_Criterio <> "", "", " AND ") & "produtos.descricao  LIKE " & var_Criterio & "", "")\r\n'
    '   End If',

    '   If optDesc.Value Then\r\n'
    '      Dim sW1 As String\r\n'
    '      Dim sW2 As String\r\n'
    '      Dim sPosEsp As Integer\r\n'
    '      If optPorPalavra.Value Then\r\n'
    '         var_Criterio = var_Criterio & IIf(var_Criterio <> "", " AND ", "") & "produtos.descricao LIKE \'%" & Replace(cboDesc.Text, "\'", "\'\'") & "%\'"\r\n'
    '      ElseIf PorPalavraDupla.Value Then\r\n'
    '         sPosEsp = InStr(Trim(cboDesc.Text), " ")\r\n'
    '         If sPosEsp > 0 Then\r\n'
    '            sW1 = Left(Trim(cboDesc.Text), sPosEsp - 1)\r\n'
    '            sW2 = Trim(Mid(cboDesc.Text, sPosEsp + 1))\r\n'
    '         Else\r\n'
    '            sW1 = Trim(cboDesc.Text)\r\n'
    '            sW2 = ""\r\n'
    '         End If\r\n'
    '         var_Criterio = var_Criterio & IIf(var_Criterio <> "", " AND ", "") & "produtos.descricao LIKE \'%" & Replace(sW1, "\'", "\'\'") & "%\'"\r\n'
    '         If Len(sW2) > 0 Then\r\n'
    '            var_Criterio = var_Criterio & " AND produtos.descricao LIKE \'%" & Replace(sW2, "\'", "\'\'") & "%\'"\r\n'
    '         End If\r\n'
    '      End If\r\n'
    '   End If',
    content)

# ── 5: Remover todas as linhas "cboConsLinha.Visible = False" dos opt*_Click ───
content = sub('remove cboConsLinha.Visible all',
    '   cboConsLinha.Visible = False\r\n',
    '',
    content, replace_all=True)

# ── 6: Substituir par chkDesc*.Visible = False por optPorPalavra/PorPalavraDupla = False
content = sub('chk Visible -> opt Visible False all',
    '   chkDescPorProduto.Visible = False\r\n'
    '   chkDescPorIniciais.Visible = False',
    '   optPorPalavra.Visible = False\r\n'
    '   PorPalavraDupla.Visible = False',
    content, replace_all=True)

# ── 7: optDesc_Click — corrigir para True (visivel quando desc marcado) ───────
content = sub('optDesc fix Visible True',
    'Private Sub optDesc_Click()\r\n'
    '   lblCategoria.Visible = False\r\n'
    '   lblDesc.Visible = True\r\n'
    '   cboDesc.Visible = True\r\n'
    '   optPorPalavra.Visible = False\r\n'
    '   PorPalavraDupla.Visible = False\r\n'
    '   lblCodBarra.Visible = False\r\n'
    '   txtCodBarra.Visible = False\r\n'
    '   cmdLocalizar.Visible = True\r\n'
    '   cboDesc.SetFocus\r\n'
    'End Sub\r\n'
    '\r\n'
    'Private Sub optTags_Click()',

    'Private Sub optDesc_Click()\r\n'
    '   lblCategoria.Visible = False\r\n'
    '   lblDesc.Visible = True\r\n'
    '   cboDesc.Visible = True\r\n'
    '   optPorPalavra.Visible = True\r\n'
    '   PorPalavraDupla.Visible = True\r\n'
    '   lblCodBarra.Visible = False\r\n'
    '   txtCodBarra.Visible = False\r\n'
    '   cmdLocalizar.Visible = True\r\n'
    '   cboDesc.SetFocus\r\n'
    'End Sub\r\n'
    '\r\n'
    'Private Sub optTags_Click()',
    content)

# ── 8: Limpar corpo de cboDesc_Click (remover refs a chk*) ───────────────────
content = sub('cboDesc_Click cleanup',
    'Private Sub cboDesc_Click()\r\n'
    '   \'If chkDescPorProduto.Value = Checked Then\r\n'
    '   \'   If cboDesc.Text = "" Then Exit Sub\r\n'
    '   \'   MostrarCriterios\r\n'
    '   \'ElseIf chkDescPorIniciais.Value = Checked Then\r\n'
    '   \'   MostrarCriterios\r\n'
    '   \'End If\r\n'
    'End Sub',

    'Private Sub cboDesc_Click()\r\n'
    'End Sub',
    content)

if errors:
    print('ERRORS:', errors)
else:
    content = content.replace('\r\n', '\n').replace('\r', '\n').replace('\n', '\r\n')
    with open(path, 'wb') as f:
        f.write(content.encode('windows-1252'))
    print('OK — arquivo gravado.')

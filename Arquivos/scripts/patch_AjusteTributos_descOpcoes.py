path = r'C:\Projeto\OnlineCommerce\Forms\Produtos_AjusteTributos.frm'
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

# ── 1: Remover definicoes dos controles chkDescPorProduto e chkDescPorIniciais
content = sub('Remove definicao chkDesc*',
    '         End\r\n'
    '         Begin VB.CheckBox chkDescPorProduto \r\n'
    '            Caption         =   "Por Produto"\r\n'
    '            ForeColor       =   &H000000FF&\r\n'
    '            Height          =   195\r\n'
    '            Left            =   2040\r\n'
    '            TabIndex        =   38\r\n'
    '            Top             =   120\r\n'
    '            Value           =   1  \'Checked\r\n'
    '            Visible         =   0   \'False\r\n'
    '            Width           =   1275\r\n'
    '         End\r\n'
    '         Begin VB.CheckBox chkDescPorIniciais \r\n'
    '            Caption         =   "Por Iniciais"\r\n'
    '            ForeColor       =   &H000000FF&\r\n'
    '            Height          =   195\r\n'
    '            Left            =   3240\r\n'
    '            TabIndex        =   37\r\n'
    '            Top             =   120\r\n'
    '            Visible         =   0   \'False\r\n'
    '            Width           =   1155\r\n'
    '         End\r\n'
    '         Begin VB.TextBox txtCodBarra ',

    '         End\r\n'
    '         Begin VB.TextBox txtCodBarra ',
    content)

# ── 2: Remover subs chkDescPorIniciais_Click e chkDescPorProduto_Click ────
content = sub('Remove chkDesc*_Click subs',
    'Private Sub chkDescPorIniciais_Click()\r\n'
    '   If optDesc.Value = Unchecked Then Exit Sub\r\n'
    '   \r\n'
    '   If chkDescPorIniciais.Value = Checked Then\r\n'
    '      cboDesc.Clear\r\n'
    '      chkDescPorProduto.Value = Unchecked\r\n'
    '      cboDesc.SetFocus\r\n'
    '   End If\r\n'
    'End Sub\r\n'
    '\r\n'
    'Private Sub chkDescPorProduto_Click()\r\n'
    '   If optDesc.Value = Unchecked Then Exit Sub\r\n'
    '   \r\n'
    '   If chkDescPorProduto.Value = Checked Then\r\n'
    '      chkDescPorIniciais.Value = Unchecked\r\n'
    '      cboDesc.SetFocus\r\n'
    '   End If\r\n'
    'End Sub\r\n'
    '\r\n'
    'Private Sub MostrarCriterios()',

    'Private Sub MostrarCriterios()',
    content)

# ── 3: Limpar cboDesc_GotFocus (nao carregar descricoes) ─────────────────
content = sub('cboDesc_GotFocus limpar',
    'Private Sub cboDesc_GotFocus()\r\n'
    '   Dim sSQL As String\r\n'
    '   Dim r As ADODB.Recordset\r\n'
    '   \r\n'
    '   If chkDescPorProduto.Value = Checked Then\r\n'
    '      cboDesc.Clear\r\n'
    '      \r\n'
    '      sSQL = "SELECT DISTINCT descricao FROM produtos ORDER BY descricao;"\r\n'
    '      Set r = dbData.OpenRecordset(sSQL)\r\n'
    '      \r\n'
    '      Do While Not r.EOF\r\n'
    '         cboDesc.AddItem ValidateNull(r("descricao"))\r\n'
    '         r.MoveNext\r\n'
    '      Loop\r\n'
    '      \r\n'
    '      If r.State <> 0 Then r.Close\r\n'
    '      Set r = Nothing\r\n'
    '      \r\n'
    '      moCombo.AttachTo cboDesc\r\n'
    '   End If\r\n'
    'End Sub',

    'Private Sub cboDesc_GotFocus()\r\n'
    'End Sub',
    content)

# ── 4: Substituir logica chk em MostrarCriterios por optPorPalavra/PorPalavraDupla
new_desc_criterio = (
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
    '   End If\r\n'
)

content = sub('MostrarCriterios nova logica desc',
    '   If chkDescPorProduto.Value = Checked Then\r\n'
    '      var_Criterio = var_Criterio & IIf(optDesc.Value, IIf(var_Criterio <> "", " AND ", "") & "produtos.descricao = \'" & cboDesc.Text & "\'", "")\r\n'
    '   ElseIf chkDescPorIniciais.Value = Checked Then\r\n'
    '      var_Criterio = Chr$(39) & cboDesc.Text & "%" & Chr(39)\r\n'
    '      var_Criterio = IIf(optDesc.Value, IIf(var_Criterio <> "", "", " AND ") & "produtos.descricao  LIKE " & var_Criterio & "", "")\r\n'
    '   End If',

    new_desc_criterio.rstrip('\r\n'),
    content)

# ── 5: Visibilidade chk → opt (replace_all=True) ─────────────────────────
content = sub('Visibilidade False chkDesc* -> opt*',
    'chkDescPorProduto.Visible = False\r\nchkDescPorIniciais.Visible = False',
    'optPorPalavra.Visible = False\r\nPorPalavraDupla.Visible = False',
    content, replace_all=True)

content = sub('Visibilidade True chkDesc* -> opt*',
    'chkDescPorProduto.Visible = True\r\nchkDescPorIniciais.Visible = True',
    'optPorPalavra.Visible = True\r\nPorPalavraDupla.Visible = True',
    content, replace_all=True)

# ── 6: cmdAtualizar.Visible = False em todos os optE*_Click (replace_all) ─
content = sub('cmdAtualizar ocultar em optE* (8x)',
    '   cmdEdicaoColetiva.Visible = True\r\n',
    '   cmdEdicaoColetiva.Visible = True\r\n'
    '   cmdAtualizar.Visible = False\r\n',
    content, replace_all=True)

# ── 7: cmdAtualizar.Visible = True apos cmdEdicaoColetiva_Click ───────────
content = sub('cmdAtualizar restaurar apos edicao',
    'cmdEdicaoColetiva.Visible = False\r\n'
    'lblEdicaoColetiva.Visible = False\r\n'
    'End Sub\r\n'
    '\r\n'
    'Private Sub txtEdicaoColetiva_LostFocus',

    'cmdEdicaoColetiva.Visible = False\r\n'
    'lblEdicaoColetiva.Visible = False\r\n'
    'cmdAtualizar.Visible = True\r\n'
    'End Sub\r\n'
    '\r\n'
    'Private Sub txtEdicaoColetiva_LostFocus',
    content)

# ── 8: UCase para optECategoria em cmdEdicaoColetiva_Click ───────────────
content = sub('UCase para optECategoria',
    'If optETags.Value Then sVal = UCase(sVal)',
    'If optETags.Value Or optECategoria.Value Then sVal = UCase(sVal)',
    content)

# ── 9: Inserir cboEdicaoColetiva_KeyPress e _Change para uppercase ────────
new_cbo_handlers = (
    'Private Sub cboEdicaoColetiva_KeyPress(KeyAscii As Integer)\r\n'
    '   If Not (optETags.Value Or optECategoria.Value) Then Exit Sub\r\n'
    '   If KeyAscii >= 97 And KeyAscii <= 122 Then\r\n'
    '      KeyAscii = KeyAscii - 32\r\n'
    '   ElseIf optETags.Value Then\r\n'
    '      If Not ((KeyAscii >= 65 And KeyAscii <= 90) Or KeyAscii = 8) Then\r\n'
    '         KeyAscii = 0\r\n'
    '      End If\r\n'
    '   End If\r\n'
    'End Sub\r\n'
    '\r\n'
    'Private Sub cboEdicaoColetiva_Change()\r\n'
    '   If Not (optETags.Value Or optECategoria.Value) Then Exit Sub\r\n'
    '   Dim pos As Integer\r\n'
    '   pos = cboEdicaoColetiva.SelStart\r\n'
    '   cboEdicaoColetiva.Text = UCase(cboEdicaoColetiva.Text)\r\n'
    '   cboEdicaoColetiva.SelStart = pos\r\n'
    'End Sub\r\n'
    '\r\n'
)

content = sub('Inserir KeyPress e Change cboEdicaoColetiva',
    'Private Sub cboEdicaoColetiva_LostFocus()',
    new_cbo_handlers + 'Private Sub cboEdicaoColetiva_LostFocus()',
    content)

if errors:
    print('ERRORS:', errors)
else:
    content = content.replace('\r\n', '\n').replace('\r', '\n').replace('\n', '\r\n')
    with open(path, 'wb') as f:
        f.write(content.encode('windows-1252'))
    print('OK — arquivo gravado.')

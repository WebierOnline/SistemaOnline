path = r'C:\Projeto\Compartilhado\Forms\Produtos_Cadastro.frm'
with open(path, 'rb') as f:
    data = f.read()
content = data.decode('windows-1252')

errors = []
R = '\r\n'

COD = 'C\xd3D. BARRA'

def sub(label, old, new, c):
    n = c.count(old)
    if n != 1:
        errors.append(f'{label}: count={n}')
        return c
    print(f'{label} OK')
    return c.replace(old, new, 1)

# ── 1: AplicarModoIS — default 1,0000 when enabling ──────────────────────────
content = sub('AplicarModoIS default fator',
    '    If iTipo = 2 Or iTipo = 3 Then' + R +
    '        txtISFatorConv.Enabled = True' + R +
    '    Else',

    '    If iTipo = 2 Or iTipo = 3 Then' + R +
    '        txtISFatorConv.Enabled = True' + R +
    '        If Trim(txtISFatorConv.Text) = "" Then txtISFatorConv.Text = "1,0000"' + R +
    '    Else',
    content)

# ── 2: cmdNovo — pre-fill 1,0000 instead of empty ───────────────────────────
content = sub('cmdNovo fator default',
    'txtISFatorConv.Text = ""' + R +
    'txtISFatorConv.Enabled = False',

    'txtISFatorConv.Text = "1,0000"' + R +
    'txtISFatorConv.Enabled = False',
    content)

# ── 3: load — if fator = 0 show 1,0000 ───────────────────────────────────────
content = sub('load fator zero to 1',
    '    txtISFatorConv.Text = Format(ValidateNull(r("fator_conversao_IS")), "##,##0.0000")' + R +
    '    AplicarModoIS',

    '    Dim dFatorIS As Double' + R +
    '    dFatorIS = CDbl(ValidateNull(r("fator_conversao_IS")))' + R +
    '    If dFatorIS = 0 Then' + R +
    '        txtISFatorConv.Text = "1,0000"' + R +
    '    Else' + R +
    '        txtISFatorConv.Text = Format(dFatorIS, "##,##0.0000")' + R +
    '    End If' + R +
    '    AplicarModoIS',
    content)

# ── 4: add txtEANCaixa_LostFocus, txtISFatorConv_LostFocus, cboTAGs_Change ───
content = sub('add LostFocus and Change events',
    'moCombo.AttachTo cboFabricante' + R +
    'End Sub' + R + R + R +
    'Private Sub cboTAGs_GotFocus()',

    'moCombo.AttachTo cboFabricante' + R +
    'End Sub' + R + R +
    'Private Sub txtEANCaixa_LostFocus()' + R +
    '    Dim s As String' + R +
    '    s = Replace(txtEANCaixa.Text, " ", "")' + R +
    '    If s = "" Then' + R +
    '        txtEANCaixa.Text = "SEM GTIN"' + R +
    '        Exit Sub' + R +
    '    End If' + R +
    '    If s = "SEM GTIN" Then Exit Sub' + R +
    '    txtEANCaixa.Text = s' + R +
    '    If Len(s) = 8 Or Len(s) = 13 Then' + R +
    '        Dim i As Integer, soma As Long, w As Integer' + R +
    '        soma = 0' + R +
    '        For i = 1 To Len(s) - 1' + R +
    '            If Len(s) = 13 Then' + R +
    '                If i Mod 2 = 1 Then w = 1 Else w = 3' + R +
    '            Else' + R +
    '                If i Mod 2 = 1 Then w = 3 Else w = 1' + R +
    '            End If' + R +
    '            soma = soma + Val(Mid(s, i, 1)) * w' + R +
    '        Next i' + R +
    '        If (10 - (soma Mod 10)) Mod 10 <> Val(Right(s, 1)) Then' + R +
    '            ShowMsg "EAN inv\xe1lido: d\xedgito verificador incorreto.", vbExclamation' + R +
    '            txtEANCaixa.SetFocus' + R +
    '        End If' + R +
    '    End If' + R +
    'End Sub' + R + R +
    'Private Sub txtISFatorConv_LostFocus()' + R +
    '    Dim sFator As String' + R +
    '    sFator = Trim(txtISFatorConv.Text)' + R +
    '    If sFator = "" Then' + R +
    '        txtISFatorConv.Text = "1,0000"' + R +
    '    Else' + R +
    '        txtISFatorConv.Text = Format(Val(Replace(Replace(sFator, ".", ""), ",", ".")), "##,##0.0000")' + R +
    '    End If' + R +
    'End Sub' + R + R +
    'Private Sub cboTAGs_Change()' + R +
    '    Dim iPos As Integer' + R +
    '    iPos = cboTAGs.SelStart' + R +
    '    cboTAGs.Text = UCase(cboTAGs.Text)' + R +
    '    cboTAGs.SelStart = iPos' + R +
    'End Sub' + R + R + R +
    'Private Sub cboTAGs_GotFocus()',
    content)

# ── 5: cboCriterios DESCRIÇÃO — marcar optPorPalavra ────────────────────────
content = sub('cboCriterios DESCRICAO optPorPalavra',
    '    optPalavrasDuplas.Visible = True' + R +
    '   cboConsProduto.SetFocus' + R +
    'ElseIf cboCriterios.Text = "FABRICANTE" Then',

    '    optPalavrasDuplas.Visible = True' + R +
    '    optPorPalavra.Value = True' + R +
    '   cboConsProduto.SetFocus' + R +
    'ElseIf cboCriterios.Text = "FABRICANTE" Then',
    content)

# ── 6a: cboConsProduto_KeyPress — digits-only for COD.BARRA / NCM ─────────
content = sub('cboConsProduto_KeyPress digits-only',
    'Private Sub cboConsProduto_KeyPress(KeyAscii As Integer)' + R +
    'KeyAscii = Asc(UCase(Chr(KeyAscii)))' + R +
    'End Sub',

    'Private Sub cboConsProduto_KeyPress(KeyAscii As Integer)' + R +
    'If cboCriterios.Text = "' + COD + '" Or cboCriterios.Text = "NCM" Then' + R +
    '    If KeyAscii <> 8 And (KeyAscii < 48 Or KeyAscii > 57) Then' + R +
    '        KeyAscii = 0' + R +
    '    End If' + R +
    'Else' + R +
    '    KeyAscii = Asc(UCase(Chr(KeyAscii)))' + R +
    'End If' + R +
    'End Sub',
    content)

# ── 6b: cboConsProduto_LostFocus — remove spaces for NCM ─────────────────
content = sub('cboConsProduto_LostFocus NCM trim',
    'If cboCriterios.Text = "' + COD + '" Then' + R +
    '    If Len(cboConsProduto) < 13 And cboConsProduto.Text <> "" Then' + R +
    '        If Len(cboConsProduto) < 6 Then' + R +
    '            cboConsProduto.Text = Format(cboConsProduto.Text, "00000")' + R +
    '        Else' + R +
    '            cboConsProduto.Text = cboConsProduto.Text' + R +
    '        End If' + R +
    '    End If' + R +
    'End If' + R +
    'End Sub' + R + R +
    'Private Sub cboCriterios_Click()',

    'If cboCriterios.Text = "' + COD + '" Then' + R +
    '    If Len(cboConsProduto) < 13 And cboConsProduto.Text <> "" Then' + R +
    '        If Len(cboConsProduto) < 6 Then' + R +
    '            cboConsProduto.Text = Format(cboConsProduto.Text, "00000")' + R +
    '        Else' + R +
    '            cboConsProduto.Text = cboConsProduto.Text' + R +
    '        End If' + R +
    '    End If' + R +
    'ElseIf cboCriterios.Text = "NCM" Then' + R +
    '    cboConsProduto.Text = Replace(cboConsProduto.Text, " ", "")' + R +
    'End If' + R +
    'End Sub' + R + R +
    'Private Sub cboCriterios_Click()',
    content)

# ── 7: cboISCST Case "00" — limpar cboISClasse ───────────────────────────────
content = sub('cboISCST Case 00 clear classe',
    '            cboISClasse.Enabled = True' + R +
    '        Case "01"',

    '            cboISClasse.Enabled = True' + R +
    '            cboISClasse.ListIndex = -1' + R +
    '            lblISCSTClass.Caption = ""' + R +
    '        Case "01"',
    content)

if errors:
    print('ERRORS:', errors)
else:
    content = content.replace('\r\n', '\n').replace('\r', '\n').replace('\n', '\r\n')
    with open(path, 'wb') as f:
        f.write(content.encode('windows-1252'))
    print('OK - arquivo gravado.')

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

CRLF = '\r\n'

# ── 1: Select Case — adicionar Case 6 (MED.) com cboEdit lista fixa ──────────
content = sub('Select Case add Case 6 MED',
    '   Case Else' + CRLF +
    '      txtEdit.Move Grid.Left + Grid.CellLeft, Grid.Top + Grid.CellTop, Grid.CellWidth, Grid.CellHeight' + CRLF +
    '      txtEdit.Text = Grid.TextMatrix(iRow, iCol)' + CRLF +
    '      txtEdit.Visible = True' + CRLF +
    '      txtEdit.SetFocus' + CRLF +
    '      txtEdit.SelStart = 0' + CRLF +
    '      txtEdit.SelLength = Len(txtEdit.Text)' + CRLF +
    'End Select',

    '   Case 6' + CRLF +
    '      cboEdit.Clear' + CRLF +
    '      With cboEdit' + CRLF +
    '         .AddItem "UN": .AddItem "CX": .AddItem "M":   .AddItem "M2"' + CRLF +
    '         .AddItem "M3": .AddItem "ML": .AddItem "KG":  .AddItem "GR"' + CRLF +
    '         .AddItem "CT": .AddItem "PO": .AddItem "SC":  .AddItem "PA"' + CRLF +
    '         .AddItem "EX": .AddItem "BJ": .AddItem "DZ":  .AddItem "PC"' + CRLF +
    '         .AddItem "DI": .AddItem "FD": .AddItem "PT"' + CRLF +
    '      End With' + CRLF +
    '      cboEdit.Move Grid.Left + Grid.CellLeft, Grid.Top + Grid.CellTop, Grid.CellWidth' + CRLF +
    '      cboEdit.Text = Grid.TextMatrix(iRow, iCol)' + CRLF +
    '      cboEdit.Visible = True' + CRLF +
    '      cboEdit.SetFocus' + CRLF +
    '   Case Else' + CRLF +
    '      txtEdit.Move Grid.Left + Grid.CellLeft, Grid.Top + Grid.CellTop, Grid.CellWidth, Grid.CellHeight' + CRLF +
    '      txtEdit.Text = Grid.TextMatrix(iRow, iCol)' + CRLF +
    '      txtEdit.Visible = True' + CRLF +
    '      txtEdit.SetFocus' + CRLF +
    '      txtEdit.SelStart = 0' + CRLF +
    '      txtEdit.SelLength = Len(txtEdit.Text)' + CRLF +
    'End Select',
    content)

# ── 2: cboEdit_KeyPress — bloquear digitacao no col 6 ────────────────────────
content = sub('cboEdit_KeyPress col 6 block',
    'Private Sub cboEdit_KeyPress(KeyAscii As Integer)' + CRLF +
    '   If iCol = 8 Then' + CRLF +
    '      If KeyAscii >= 97 And KeyAscii <= 122 Then KeyAscii = KeyAscii - 32' + CRLF +
    '   End If' + CRLF +
    'End Sub',

    'Private Sub cboEdit_KeyPress(KeyAscii As Integer)' + CRLF +
    '   If iCol = 6 Then' + CRLF +
    '      KeyAscii = 0' + CRLF +
    '   ElseIf iCol = 8 Then' + CRLF +
    '      If KeyAscii >= 97 And KeyAscii <= 122 Then KeyAscii = KeyAscii - 32' + CRLF +
    '   End If' + CRLF +
    'End Sub',
    content)

# ── 3: cboEdit_LostFocus — validar col 6 contra lista fixa ───────────────────
content = sub('cboEdit_LostFocus col 6 validate',
    'Private Sub cboEdit_LostFocus()' + CRLF +
    '   If iRow > 0 Then' + CRLF +
    '      Dim sVal As String' + CRLF +
    '      sVal = Trim(cboEdit.Text)' + CRLF +
    '      If iCol = 8 Then sVal = UCase(sVal)' + CRLF +
    '      Grid.TextMatrix(iRow, iCol) = sVal' + CRLF +
    '   End If' + CRLF +
    '   cboEdit.Visible = False' + CRLF +
    'End Sub',

    'Private Sub cboEdit_LostFocus()' + CRLF +
    '   If iRow > 0 Then' + CRLF +
    '      Dim sVal As String' + CRLF +
    '      sVal = Trim(cboEdit.Text)' + CRLF +
    '      If iCol = 8 Then sVal = UCase(sVal)' + CRLF +
    '      If iCol = 6 Then' + CRLF +
    '         Dim sMed As String' + CRLF +
    '         sMed = ",UN,CX,M,M2,M3,ML,KG,GR,CT,PO,SC,PA,EX,BJ,DZ,PC,DI,FD,PT,"' + CRLF +
    '         If InStr(sMed, "," & UCase(sVal) & ",") = 0 Then sVal = Grid.TextMatrix(iRow, iCol)' + CRLF +
    '      End If' + CRLF +
    '      Grid.TextMatrix(iRow, iCol) = sVal' + CRLF +
    '   End If' + CRLF +
    '   cboEdit.Visible = False' + CRLF +
    'End Sub',
    content)

# ── 4: txtEdit_LostFocus — col 3 remove espacos, col 4/5 trim ────────────────
content = sub('txtEdit_LostFocus regras',
    'Private Sub txtEdit_LostFocus()' + CRLF +
    "'criado por mim" + CRLF +
    'If iCol = 6 Then' + CRLF +
    '    txtEdit.Text = Replace(txtEdit.Text, ".", "")' + CRLF +
    '    txtEdit.Text = Trim(txtEdit.Text)' + CRLF +
    'End If' + CRLF +
    CRLF +
    'Grid.TextMatrix(iRow, iCol) = IIf(txtEdit.Text = "", 0, txtEdit.Text)' + CRLF +
    'txtEdit.Visible = False' + CRLF +
    'End Sub',

    'Private Sub txtEdit_LostFocus()' + CRLF +
    'If iCol = 3 Then' + CRLF +
    '   txtEdit.Text = Replace(txtEdit.Text, " ", "")' + CRLF +
    'ElseIf iCol = 4 Or iCol = 5 Then' + CRLF +
    '   txtEdit.Text = Trim(txtEdit.Text)' + CRLF +
    'End If' + CRLF +
    CRLF +
    'If iCol >= 3 And iCol <= 5 Then' + CRLF +
    '   Grid.TextMatrix(iRow, iCol) = txtEdit.Text' + CRLF +
    'Else' + CRLF +
    '   Grid.TextMatrix(iRow, iCol) = IIf(txtEdit.Text = "", 0, txtEdit.Text)' + CRLF +
    'End If' + CRLF +
    'txtEdit.Visible = False' + CRLF +
    'End Sub',
    content)

# ── 5: Adicionar txtEdit_KeyPress e txtEdit_Change antes de txtEdit_GotFocus ──
content = sub('add txtEdit_KeyPress e Change',
    'Private Sub txtEdit_GotFocus()',

    'Private Sub txtEdit_KeyPress(KeyAscii As Integer)' + CRLF +
    '   Select Case iCol' + CRLF +
    '   Case 3' + CRLF +
    '      If KeyAscii <> 8 Then' + CRLF +
    '         If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0' + CRLF +
    '      End If' + CRLF +
    '   Case 4, 5' + CRLF +
    '      If KeyAscii = 8 Then Exit Sub' + CRLF +
    '      If KeyAscii >= 97 And KeyAscii <= 122 Then KeyAscii = KeyAscii - 32: Exit Sub' + CRLF +
    '      If (KeyAscii >= 65 And KeyAscii <= 90) Then Exit Sub' + CRLF +
    '      If KeyAscii = 32 Then Exit Sub' + CRLF +
    '      If KeyAscii >= 48 And KeyAscii <= 57 Then Exit Sub' + CRLF +
    '      If KeyAscii = 45 Or KeyAscii = 46 Or KeyAscii = 47 Then Exit Sub' + CRLF +
    '      KeyAscii = 0' + CRLF +
    '   End Select' + CRLF +
    'End Sub' + CRLF +
    CRLF +
    'Private Sub txtEdit_Change()' + CRLF +
    '   If iCol = 4 Or iCol = 5 Then' + CRLF +
    '      Dim p As Integer' + CRLF +
    '      p = txtEdit.SelStart' + CRLF +
    '      txtEdit.Text = UCase(txtEdit.Text)' + CRLF +
    '      txtEdit.SelStart = p' + CRLF +
    '   End If' + CRLF +
    'End Sub' + CRLF +
    CRLF +
    'Private Sub txtEdit_GotFocus()',
    content)

if errors:
    print('ERRORS:', errors)
else:
    content = content.replace('\r\n', '\n').replace('\r', '\n').replace('\n', '\r\n')
    with open(path, 'wb') as f:
        f.write(content.encode('windows-1252'))
    print('OK - arquivo gravado.')

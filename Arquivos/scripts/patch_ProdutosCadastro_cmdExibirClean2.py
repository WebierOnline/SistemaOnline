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

# ── 1: cboConsProduto_LostFocus — strip espaços (CÓD.BARRA) e não-dígitos (NCM) antes dos outros tratamentos
content = sub('LostFocus strip CodBarra e NCM',
    'If cboCriterios.Text = "' + COD + '" Then' + R +
    '    If Len(cboConsProduto) < 13 And cboConsProduto.Text <> "" Then' + R +
    '        If Len(cboConsProduto) < 6 Then' + R +
    '            cboConsProduto.Text = Format(cboConsProduto.Text, "00000")' + R +
    '        Else' + R +
    '            cboConsProduto.Text = cboConsProduto.Text' + R +
    '        End If' + R +
    '    End If' + R +
    'ElseIf cboCriterios.Text = "NCM" Then' + R +
    '    cboConsProduto.Text = Replace(cboConsProduto.Text, " ", "")',

    'If cboCriterios.Text = "' + COD + '" Then' + R +
    '    cboConsProduto.Text = Replace(Replace(Trim(cboConsProduto.Text), Chr(160), ""), " ", "")' + R +
    '    If Len(cboConsProduto) < 13 And cboConsProduto.Text <> "" Then' + R +
    '        If Len(cboConsProduto) < 6 Then' + R +
    '            cboConsProduto.Text = Format(cboConsProduto.Text, "00000")' + R +
    '        Else' + R +
    '            cboConsProduto.Text = cboConsProduto.Text' + R +
    '        End If' + R +
    '    End If' + R +
    'ElseIf cboCriterios.Text = "NCM" Then' + R +
    '    Dim sLFncm As String, iLFncm As Integer' + R +
    '    sLFncm = ""' + R +
    '    For iLFncm = 1 To Len(cboConsProduto.Text)' + R +
    '        If Mid(cboConsProduto.Text, iLFncm, 1) >= "0" And Mid(cboConsProduto.Text, iLFncm, 1) <= "9" Then' + R +
    '            sLFncm = sLFncm & Mid(cboConsProduto.Text, iLFncm, 1)' + R +
    '        End If' + R +
    '    Next iLFncm' + R +
    '    cboConsProduto.Text = sLFncm',
    content)

# ── 2: cmdExibir_Click — substituir bloco que escrevia em .Text por variável local sConsProd
# (escrever em cboConsProduto.Text dispara Change que chama cmdExibir_Click recursivamente)
content = sub('cmdExibir sConsProd local var',
    '   If cboCriterios.Text = "' + COD + '" Then' + R +
    '       cboConsProduto.Text = Replace(cboConsProduto.Text, " ", "")' + R +
    '   ElseIf cboCriterios.Text = "NCM" Then' + R +
    '       Dim sNCM As String, iNCM As Integer' + R +
    '       sNCM = ""' + R +
    '       For iNCM = 1 To Len(cboConsProduto.Text)' + R +
    '           If Mid(cboConsProduto.Text, iNCM, 1) >= "0" And Mid(cboConsProduto.Text, iNCM, 1) <= "9" Then' + R +
    '               sNCM = sNCM & Mid(cboConsProduto.Text, iNCM, 1)' + R +
    '           End If' + R +
    '       Next iNCM' + R +
    '       cboConsProduto.Text = sNCM' + R +
    '   End If',

    '   Dim sConsProd As String' + R +
    '   sConsProd = Replace(Replace(Trim(cboConsProduto.Text), Chr(160), ""), " ", "")' + R +
    '   If cboCriterios.Text = "NCM" Then' + R +
    '       Dim sNCMfilt As String, iNCMfilt As Integer' + R +
    '       sNCMfilt = ""' + R +
    '       For iNCMfilt = 1 To Len(sConsProd)' + R +
    '           If Mid(sConsProd, iNCMfilt, 1) >= "0" And Mid(sConsProd, iNCMfilt, 1) <= "9" Then' + R +
    '               sNCMfilt = sNCMfilt & Mid(sConsProd, iNCMfilt, 1)' + R +
    '           End If' + R +
    '       Next iNCMfilt' + R +
    '       sConsProd = sNCMfilt' + R +
    '   End If',
    content)

# ── 3: SQL CÓD. BARRA — usar sConsProd em vez de cboConsProduto.Text
content = sub('SQL CodBarra usa sConsProd',
    '      sSQL = sSQL & "(produtos.cod_barra = \'" & cboConsProduto.Text & "\') AND (produtos.codigo <> 1) ORDER BY " & INDICE',
    '      sSQL = sSQL & "(produtos.cod_barra = \'" & sConsProd & "\') AND (produtos.codigo <> 1) ORDER BY " & INDICE',
    content)

# ── 4: SQL NCM — usar sConsProd em vez de cboConsProduto.Text
content = sub('SQL NCM usa sConsProd',
    '      sSQL = sSQL & "(produtos.NCM = \'" & cboConsProduto.Text & "\') AND (produtos.codigo <> 1) ORDER BY " & INDICE',
    '      sSQL = sSQL & "(produtos.NCM = \'" & sConsProd & "\') AND (produtos.codigo <> 1) ORDER BY " & INDICE',
    content)

if errors:
    print('ERRORS:', errors)
else:
    content = content.replace('\r\n', '\n').replace('\r', '\n').replace('\n', '\r\n')
    with open(path, 'wb') as f:
        f.write(content.encode('windows-1252'))
    print('OK - arquivo gravado.')

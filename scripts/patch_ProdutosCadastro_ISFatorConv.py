path = r'C:\Projeto\Compartilhado\Forms\Produtos_Cadastro.frm'
with open(path, 'rb') as f:
    data = f.read()
content = data.decode('windows-1252')

errors = []
R = '\r\n'

def sub(label, old, new, c):
    n = c.count(old)
    if n != 1:
        errors.append(f'{label}: count={n}')
        return c
    print(f'{label} OK')
    return c.replace(old, new, 1)

# helper locale-safe para SQL (trata vazio como 0)
fatorSQL = 'Replace(CStr(CDbl(IIf(Trim(txtISFatorConv.Text) = "", "0", txtISFatorConv.Text))), ",", ".")'

# ── 1: AplicarModoIS — habilitar/desabilitar txtISFatorConv ───────────────────
content = sub('AplicarModoIS enable/disable',
    'Private Sub AplicarModoIS()' + R +
    '    AplicarModoIBSCBS' + R +
    'End Sub',

    'Private Sub AplicarModoIS()' + R +
    '    AplicarModoIBSCBS' + R +
    '    Dim iTipo As Integer' + R +
    '    iTipo = Val(Left(cboTipoIS.Text & "0", 1))' + R +
    '    If iTipo = 2 Or iTipo = 3 Then' + R +
    '        txtISFatorConv.Enabled = True' + R +
    '    Else' + R +
    '        txtISFatorConv.Enabled = False' + R +
    '        txtISFatorConv.Text = ""' + R +
    '    End If' + R +
    'End Sub',
    content)

# ── 2: cboTipoIS_Click — acionar AplicarModoIS ao mudar tipo ─────────────────
content = sub('cboTipoIS_Click novo evento',
    'Private Sub AplicarModoIBSCBS()',

    'Private Sub cboTipoIS_Click()' + R +
    '    AplicarModoIS' + R +
    'End Sub' + R +
    R +
    'Private Sub AplicarModoIBSCBS()',
    content)

# ── 3: Carregar fator_conversao_IS ao abrir produto ───────────────────────────
content = sub('load fator_conversao_IS',
    'If Not IsNull(r("tipo_calculo_is")) Then SelecionarNoCombo cboTipoIS, CStr(r("tipo_calculo_is")), True',

    'If Not IsNull(r("tipo_calculo_is")) Then SelecionarNoCombo cboTipoIS, CStr(r("tipo_calculo_is")), True' + R +
    '    txtISFatorConv.Text = Format(IIf(IsNull(r("fator_conversao_IS")), 0, CDbl(r("fator_conversao_IS"))), "##,##0.0000")' + R +
    '    AplicarModoIS',
    content)

# ── 4: cmdNovo — limpar e desabilitar txtISFatorConv ─────────────────────────
content = sub('cmdNovo clear fator',
    'cboISClasse.ListIndex = -1' + R +
    'cboTipoIS.ListIndex = -1',

    'cboISClasse.ListIndex = -1' + R +
    'cboTipoIS.ListIndex = -1' + R +
    'txtISFatorConv.Text = ""' + R +
    'txtISFatorConv.Enabled = False',
    content)

# ── 5a: INSERT — coluna fator_conversao_IS apos tipo_calculo_is ───────────────
content = sub('INSERT col fator_conversao_IS',
    'cClassTrib_IS, tipo_calculo_is, TAGS,',
    'cClassTrib_IS, tipo_calculo_is, fator_conversao_IS, TAGS,',
    content)

# ── 5b: INSERT — valor de fator_conversao_IS apos tipo_calculo_is ────────────
content = sub('INSERT val fator_conversao_IS',
    '& IIf(cboTipoIS.ListIndex >= 0, Val(Left(cboTipoIS.Text, 1)), "NULL") & ", " & IIf(Trim(cboTAGs.Text) = "", "NULL"',
    '& IIf(cboTipoIS.ListIndex >= 0, Val(Left(cboTipoIS.Text, 1)), "NULL") & ", " & ' + fatorSQL + ' & ", " & IIf(Trim(cboTAGs.Text) = "", "NULL"',
    content)

# ── 6: UPDATE — campo fator_conversao_IS antes de TAGS ───────────────────────
content = sub('UPDATE fator_conversao_IS',
    'sSQL = sSQL & "tipo_calculo_is = " & IIf(cboTipoIS.ListIndex >= 0, Val(Left(cboTipoIS.Text, 1)), "NULL") & ", "' + R +
    '    sSQL = sSQL & "TAGS = "',

    'sSQL = sSQL & "tipo_calculo_is = " & IIf(cboTipoIS.ListIndex >= 0, Val(Left(cboTipoIS.Text, 1)), "NULL") & ", "' + R +
    '    sSQL = sSQL & "fator_conversao_IS = " & ' + fatorSQL + ' & ", "' + R +
    '    sSQL = sSQL & "TAGS = "',
    content)

if errors:
    print('ERRORS:', errors)
else:
    content = content.replace('\r\n', '\n').replace('\r', '\n').replace('\n', '\r\n')
    with open(path, 'wb') as f:
        f.write(content.encode('windows-1252'))
    print('OK - arquivo gravado.')

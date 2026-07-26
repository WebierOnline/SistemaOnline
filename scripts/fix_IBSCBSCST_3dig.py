# -*- coding: windows-1252 -*-
path = 'C:/Projeto/Compartilhado/Forms/Produtos_Cadastro.frm'
f = open(path, 'rb')
raw = f.read()
f.close()
text = raw.decode('windows-1252')

errors = []
def apply(text, old, new, tag):
    cnt = text.count(old)
    if cnt == 0:
        print('JA APLICADO', tag)
        return text
    if cnt != 1:
        errors.append('ERRO {}: {}x'.format(tag, cnt))
        return text
    print('OK', tag)
    return text.replace(old, new)

# 1. PreencheIBSCBS - novos itens com 3 digitos
text = apply(text,
    'cboIBSCBSCST.AddItem "00 - N\xe3o Incidente"\r\n'
    'cboIBSCBSCST.AddItem "01 - Tributado Integralmente"\r\n'
    'cboIBSCBSCST.AddItem "02 - Al\xedquota Reduzida"\r\n'
    'cboIBSCBSCST.AddItem "10 - Imune / Isento"\r\n'
    'cboIBSCBSCST.AddItem "90 - Outros"',
    'cboIBSCBSCST.AddItem "000 - Tributa\xe7\xe3o integral"\r\n'
    'cboIBSCBSCST.AddItem "010 - Al\xedquotas uniformes"\r\n'
    'cboIBSCBSCST.AddItem "200 - Al\xedquota reduzida"\r\n'
    'cboIBSCBSCST.AddItem "400 - Isen\xe7\xe3o"\r\n'
    'cboIBSCBSCST.AddItem "620 - Tributa\xe7\xe3o Monof\xe1sica"\r\n'
    'cboIBSCBSCST.AddItem "900 - Outros"',
    '1-PreencheIBSCBS')

# 2. MostrarDados_Produto - Left 2 -> 3
text = apply(text,
    'SelecionarNoCombo cboIBSCBSCST, Left(ValidateNull(r("IBSCBSCST")) & "  ", 2), True',
    'SelecionarNoCombo cboIBSCBSCST, Left(ValidateNull(r("IBSCBSCST")) & "   ", 3), True',
    '2-MostrarDados')

# 3. cboCST_Click Simples
text = apply(text,
    '        \' --- REFORMA TRIBUT\xc1RIA (CSTs INICIAIS) ---\r\n'
    '        Select Case cboCST.Text\r\n'
    '            Case "101", "102", "103", "400", "900"\r\n'
    '                SelecionarNoCombo cboIBSCBSCST, "01", True\r\n'
    '            Case "500", "201", "202", "203"\r\n'
    '                SelecionarNoCombo cboIBSCBSCST, "02", True\r\n'
    '        End Select',
    '        \' --- REFORMA TRIBUT\xc1RIA (CSTs INICIAIS) ---\r\n'
    '        Select Case cboCST.Text\r\n'
    '            Case "101", "102", "103", "400", "900"\r\n'
    '                SelecionarNoCombo cboIBSCBSCST, "000", True\r\n'
    '            Case "500", "201", "202", "203"\r\n'
    '                SelecionarNoCombo cboIBSCBSCST, "620", True\r\n'
    '        End Select',
    '3-cboCST-Simples')

# 4. cboCST_Click Normal
text = apply(text,
    '        \' --- REFORMA TRIBUT\xc1RIA (CSTs INICIAIS) ---\r\n'
    '        Select Case Left(cboCST.Text, 2)\r\n'
    '            Case "00", "20": SelecionarNoCombo cboIBSCBSCST, "01", True\r\n'
    '            Case "10", "60", "70": SelecionarNoCombo cboIBSCBSCST, "02", True\r\n'
    '            Case Else: SelecionarNoCombo cboIBSCBSCST, "01", True\r\n'
    '        End Select',
    '        \' --- REFORMA TRIBUT\xc1RIA (CSTs INICIAIS) ---\r\n'
    '        Select Case Left(cboCST.Text, 2)\r\n'
    '            Case "00", "20": SelecionarNoCombo cboIBSCBSCST, "000", True\r\n'
    '            Case "60": SelecionarNoCombo cboIBSCBSCST, "620", True\r\n'
    '            Case "10", "70": SelecionarNoCombo cboIBSCBSCST, "200", True\r\n'
    '            Case Else: SelecionarNoCombo cboIBSCBSCST, "000", True\r\n'
    '        End Select',
    '4-cboCST-Normal')

# 5a. AtualizarReforma - isencao
text = apply(text,
    '        \' --- IBS/CBS 00 + PIS/COFINS 06: al\xedquota zero (essenciais) ---\r\n'
    '        Case "ALIMENTOS (CESTA B\xc1SICA)", "HORTIFR\xdaTI", "CARNES", "LATIC\xcdNIOS", "CONGELADOS"\r\n'
    '            sIBSCST = "00"\r\n',
    '        \' --- IBS/CBS 400 + PIS/COFINS 06: isen\xe7\xe3o (essenciais) ---\r\n'
    '        Case "ALIMENTOS (CESTA B\xc1SICA)", "HORTIFR\xdaTI", "CARNES", "LATIC\xcdNIOS", "CONGELADOS"\r\n'
    '            sIBSCST = "400"\r\n',
    '5a-isencao')

# 5b. AtualizarReforma - monofasico
text = apply(text,
    '        \' --- IBS/CBS 01 + PIS/COFINS 04: monof\xe1sico ---\r\n'
    '        Case "HIGIENE E PERFUMARIA", "BEBIDAS", "BEBIDAS (A\xc7UCARADAS)", "BEBIDAS (ALCO\xd3LICAS)", "TABACARIA"\r\n'
    '            sIBSCST = "01"\r\n',
    '        \' --- IBS/CBS 620 + PIS/COFINS 04: monof\xe1sico ---\r\n'
    '        Case "HIGIENE E PERFUMARIA", "BEBIDAS", "BEBIDAS (A\xc7UCARADAS)", "BEBIDAS (ALCO\xd3LICAS)", "TABACARIA"\r\n'
    '            sIBSCST = "620"\r\n',
    '5b-monofasico')

# 5c. AtualizarReforma - reduzida
text = apply(text,
    '        \' --- IBS/CBS 02 + PIS/COFINS 01: al\xedquota reduzida (alimentos processados) ---\r\n'
    '        Case "ALIMENTOS", "FRIOS E LATIC\xcdNIOS", "PADARIA E CONFEITARIA"\r\n'
    '            sIBSCST = "02"\r\n',
    '        \' --- IBS/CBS 200 + PIS/COFINS 01: al\xedquota reduzida (alimentos processados) ---\r\n'
    '        Case "ALIMENTOS", "FRIOS E LATIC\xcdNIOS", "PADARIA E CONFEITARIA"\r\n'
    '            sIBSCST = "200"\r\n',
    '5c-reduzida')

# 5d. AtualizarReforma - integral
text = apply(text,
    '        \' --- IBS/CBS 01 + PIS/COFINS 01: tributa\xe7\xe3o integral (limpeza, bazar, pet shop, etc.) ---\r\n'
    '        Case Else\r\n'
    '            sIBSCST = "01"\r\n',
    '        \' --- IBS/CBS 000 + PIS/COFINS 01: tributa\xe7\xe3o integral (limpeza, bazar, pet shop, etc.) ---\r\n'
    '        Case Else\r\n'
    '            sIBSCST = "000"\r\n',
    '5d-integral')

# 6. cboIBSCBSCST_Click - Left 2->3, novos cases
text = apply(text,
    'Private Sub cboIBSCBSCST_Click()\r\n'
    '    Dim sCST As String\r\n'
    '    sCST = Left(cboIBSCBSCST.Text, 2)\r\n'
    '    Select Case sCST\r\n'
    '        Case "00", "10"  \' N\xe3o Incidente / Imune\r\n'
    '            txtCBSpAliq.Text = "0,00"\r\n'
    '            txtIBSUFpAliq.Text = "0,00"\r\n'
    '            txtIBSMunpAliq.Text = "0,00"\r\n'
    '            txtISpIS.Text = "0,00"\r\n'
    '            SelecionarNoCombo cboISCST, "00", True\r\n'
    '        Case "02"  \' Al\xedquota Reduzida\r\n'
    '            txtCBSpAliq.Text = "3,50"\r\n'
    '            txtIBSUFpAliq.Text = "7,10"\r\n'
    '            txtIBSMunpAliq.Text = "0,00"\r\n'
    '            txtISpIS.Text = "0,00"\r\n'
    '            SelecionarNoCombo cboISCST, "00", True\r\n'
    '        Case Else  \' 01 Integral / 90 Outros\r\n'
    '            txtCBSpAliq.Text = "8,80"\r\n'
    '            txtIBSUFpAliq.Text = "17,70"\r\n'
    '            txtIBSMunpAliq.Text = "0,00"\r\n'
    '    End Select\r\n'
    '    AtualizarCargaReforma\r\n'
    'End Sub',
    'Private Sub cboIBSCBSCST_Click()\r\n'
    '    Dim sCST As String\r\n'
    '    sCST = Left(cboIBSCBSCST.Text, 3)\r\n'
    '    Select Case sCST\r\n'
    '        Case "400"  \' Isen\xe7\xe3o\r\n'
    '            txtCBSpAliq.Text = "0,00"\r\n'
    '            txtIBSUFpAliq.Text = "0,00"\r\n'
    '            txtIBSMunpAliq.Text = "0,00"\r\n'
    '            txtISpIS.Text = "0,00"\r\n'
    '            SelecionarNoCombo cboISCST, "00", True\r\n'
    '        Case "200"  \' Al\xedquota Reduzida\r\n'
    '            txtCBSpAliq.Text = "3,50"\r\n'
    '            txtIBSUFpAliq.Text = "7,10"\r\n'
    '            txtIBSMunpAliq.Text = "0,00"\r\n'
    '            txtISpIS.Text = "0,00"\r\n'
    '            SelecionarNoCombo cboISCST, "00", True\r\n'
    '        Case "620"  \' Tributa\xe7\xe3o Monof\xe1sica\r\n'
    '            txtCBSpAliq.Text = "8,80"\r\n'
    '            txtIBSUFpAliq.Text = "17,70"\r\n'
    '            txtIBSMunpAliq.Text = "0,00"\r\n'
    '            txtISpIS.Text = "0,00"\r\n'
    '            SelecionarNoCombo cboISCST, "00", True\r\n'
    '        Case Else  \' 000 Integral / 010 Uniformes / 900 Outros\r\n'
    '            txtCBSpAliq.Text = "8,80"\r\n'
    '            txtIBSUFpAliq.Text = "17,70"\r\n'
    '            txtIBSMunpAliq.Text = "0,00"\r\n'
    '    End Select\r\n'
    '    AtualizarCargaReforma\r\n'
    'End Sub',
    '6-cboIBSCBSCST_Click')

# 7. INSERT SQL - Left 2 -> 3
text = apply(text,
    "& \"', '\" & Left(cboIBSCBSCST.Text, 2) & \"', \"",
    "& \"', '\" & Left(cboIBSCBSCST.Text, 3) & \"', \"",
    '7-INSERT')

# 8. UPDATE SQL - Left 2 -> 3
text = apply(text,
    '    sSQL = sSQL & "IBSCBSCST = \'" & Left(cboIBSCBSCST.Text, 2) & "\', "',
    '    sSQL = sSQL & "IBSCBSCST = \'" & Left(cboIBSCBSCST.Text, 3) & "\', "',
    '8-UPDATE')

for e in errors:
    print(e)
if not errors:
    out = text.encode('windows-1252')
    out = out.replace(b'\r\n', b'\n').replace(b'\r', b'\n').replace(b'\n', b'\r\n')
    open(path, 'wb').write(out)
    print('Arquivo salvo.')

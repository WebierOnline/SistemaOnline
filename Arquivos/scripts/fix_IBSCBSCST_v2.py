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

# 1. PreencheIBSCBS - 11 novos itens
text = apply(text,
    'cboIBSCBSCST.AddItem "000 - Tributa\xe7\xe3o integral"\r\n'
    'cboIBSCBSCST.AddItem "010 - Al\xedquotas uniformes"\r\n'
    'cboIBSCBSCST.AddItem "200 - Al\xedquota reduzida"\r\n'
    'cboIBSCBSCST.AddItem "400 - Isen\xe7\xe3o"\r\n'
    'cboIBSCBSCST.AddItem "620 - Tributa\xe7\xe3o Monof\xe1sica"\r\n'
    'cboIBSCBSCST.AddItem "900 - Outros"',
    'cboIBSCBSCST.AddItem "000 - Tributa\xe7\xe3o Integral"\r\n'
    'cboIBSCBSCST.AddItem "010 - Tributa\xe7\xe3o com Al\xedquotas Uniformes"\r\n'
    'cboIBSCBSCST.AddItem "011 - Al\xedquotas Uniformes Reduzidas"\r\n'
    'cboIBSCBSCST.AddItem "200 - Al\xedquota Reduzida"\r\n'
    'cboIBSCBSCST.AddItem "220 - Al\xedquota Fixa"\r\n'
    'cboIBSCBSCST.AddItem "400 - Isen\xe7\xe3o"\r\n'
    'cboIBSCBSCST.AddItem "410 - Imunidade e N\xe3o Incid\xeancia"\r\n'
    'cboIBSCBSCST.AddItem "510 - Diferimento"\r\n'
    'cboIBSCBSCST.AddItem "550 - Suspens\xe3o"\r\n'
    'cboIBSCBSCST.AddItem "600 - Tributa\xe7\xe3o Monof\xe1sica"\r\n'
    'cboIBSCBSCST.AddItem "820 - Tributa\xe7\xe3o em documento espec\xedfico"',
    '1-PreencheIBSCBS')

# 2. cboIBSCBSCST_Click - novos cases + lblIBSCBSCST
text = apply(text,
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
    'Private Sub cboIBSCBSCST_Click()\r\n'
    '    Dim sCST As String\r\n'
    '    sCST = Left(cboIBSCBSCST.Text, 3)\r\n'
    '    Select Case sCST\r\n'
    '        Case "400"  \' Isen\xe7\xe3o\r\n'
    '            txtCBSpAliq.Text = "0,00": txtIBSUFpAliq.Text = "0,00"\r\n'
    '            txtIBSMunpAliq.Text = "0,00": txtISpIS.Text = "0,00"\r\n'
    '            SelecionarNoCombo cboISCST, "00", True\r\n'
    '            lblIBSCBSCST.Caption = "Transporte p\xfablico coletivo e casos isentos por lei."\r\n'
    '        Case "410"  \' Imunidade e N\xe3o Incid\xeancia\r\n'
    '            txtCBSpAliq.Text = "0,00": txtIBSUFpAliq.Text = "0,00"\r\n'
    '            txtIBSMunpAliq.Text = "0,00": txtISpIS.Text = "0,00"\r\n'
    '            SelecionarNoCombo cboISCST, "00", True\r\n'
    '            lblIBSCBSCST.Caption = "Exporta\xe7\xf5es e entidades imunes."\r\n'
    '        Case "510"  \' Diferimento\r\n'
    '            txtCBSpAliq.Text = "0,00": txtIBSUFpAliq.Text = "0,00"\r\n'
    '            txtIBSMunpAliq.Text = "0,00": txtISpIS.Text = "0,00"\r\n'
    '            SelecionarNoCombo cboISCST, "00", True\r\n'
    '            lblIBSCBSCST.Caption = "Energia el\xe9trica e casos onde o imposto \xe9 pago depois."\r\n'
    '        Case "550"  \' Suspens\xe3o\r\n'
    '            txtCBSpAliq.Text = "0,00": txtIBSUFpAliq.Text = "0,00"\r\n'
    '            txtIBSMunpAliq.Text = "0,00": txtISpIS.Text = "0,00"\r\n'
    '            SelecionarNoCombo cboISCST, "00", True\r\n'
    '            lblIBSCBSCST.Caption = "Regimes especiais (ZFM, REIDI, REPORTO)."\r\n'
    '        Case "200"  \' Al\xedquota Reduzida\r\n'
    '            txtCBSpAliq.Text = "3,50": txtIBSUFpAliq.Text = "7,10"\r\n'
    '            txtIBSMunpAliq.Text = "0,00": txtISpIS.Text = "0,00"\r\n'
    '            SelecionarNoCombo cboISCST, "00", True\r\n'
    '            lblIBSCBSCST.Caption = "Alimentos, sa\xfade, educa\xe7\xe3o (redu\xe7\xf5es de 30%, 60% ou 100%)."\r\n'
    '        Case "220"  \' Al\xedquota Fixa\r\n'
    '            txtCBSpAliq.Text = "3,50": txtIBSUFpAliq.Text = "7,10"\r\n'
    '            txtIBSMunpAliq.Text = "0,00": txtISpIS.Text = "0,00"\r\n'
    '            SelecionarNoCombo cboISCST, "00", True\r\n'
    '            lblIBSCBSCST.Caption = "Incorpora\xe7\xf5es imobili\xe1rias e parcelamento de solo."\r\n'
    '        Case "011"  \' Al\xedquotas Uniformes Reduzidas\r\n'
    '            txtCBSpAliq.Text = "3,50": txtIBSUFpAliq.Text = "7,10"\r\n'
    '            txtIBSMunpAliq.Text = "0,00": txtISpIS.Text = "0,00"\r\n'
    '            SelecionarNoCombo cboISCST, "00", True\r\n'
    '            lblIBSCBSCST.Caption = "Planos de sa\xfade, funer\xe1rias e loterias."\r\n'
    '        Case "600"  \' Tributa\xe7\xe3o Monof\xe1sica\r\n'
    '            txtCBSpAliq.Text = "8,80": txtIBSUFpAliq.Text = "17,70"\r\n'
    '            txtIBSMunpAliq.Text = "0,00": txtISpIS.Text = "0,00"\r\n'
    '            SelecionarNoCombo cboISCST, "00", True\r\n'
    '            lblIBSCBSCST.Caption = "Combust\xedveis (incide uma \xfanica vez)."\r\n'
    '        Case "010"  \' Tributa\xe7\xe3o com Al\xedquotas Uniformes\r\n'
    '            txtCBSpAliq.Text = "8,80": txtIBSUFpAliq.Text = "17,70"\r\n'
    '            txtIBSMunpAliq.Text = "0,00"\r\n'
    '            lblIBSCBSCST.Caption = "Setor financeiro e opera\xe7\xf5es espec\xedficas."\r\n'
    '        Case "820"  \' Tributa\xe7\xe3o em documento espec\xedfico\r\n'
    '            txtCBSpAliq.Text = "8,80": txtIBSUFpAliq.Text = "17,70"\r\n'
    '            txtIBSMunpAliq.Text = "0,00"\r\n'
    '            lblIBSCBSCST.Caption = "Quando o imposto \xe9 apurado fora da NF-e principal."\r\n'
    '        Case Else  \' 000 - Tributa\xe7\xe3o Integral\r\n'
    '            txtCBSpAliq.Text = "8,80": txtIBSUFpAliq.Text = "17,70"\r\n'
    '            txtIBSMunpAliq.Text = "0,00"\r\n'
    '            lblIBSCBSCST.Caption = "Al\xedquota cheia (padr\xe3o) sem benef\xedcios."\r\n'
    '    End Select\r\n'
    '    AtualizarCargaReforma\r\n'
    'End Sub',
    '2-cboIBSCBSCST_Click')

# 3. AtualizarReforma: "620" -> "600"
text = apply(text,
    '            sIBSCST = "620"\r\n',
    '            sIBSCST = "600"\r\n',
    '3-AtualizarReforma-620->600')

# 4a. cboCST_Click Simples: "620" -> "600"
text = apply(text,
    '                SelecionarNoCombo cboIBSCBSCST, "620", True\r\n'
    '        End Select\r\n'
    '\r\n'
    '    Else',
    '                SelecionarNoCombo cboIBSCBSCST, "600", True\r\n'
    '        End Select\r\n'
    '\r\n'
    '    Else',
    '4a-cboCST-Simples-620->600')

# 4b. cboCST_Click Normal: "620" -> "600"
text = apply(text,
    '            Case "60": SelecionarNoCombo cboIBSCBSCST, "620", True\r\n',
    '            Case "60": SelecionarNoCombo cboIBSCBSCST, "600", True\r\n',
    '4b-cboCST-Normal-620->600')

for e in errors: print(e)
if not errors:
    out = text.encode('windows-1252')
    out = out.replace(b'\r\n', b'\n').replace(b'\r', b'\n').replace(b'\n', b'\r\n')
    open(path, 'wb').write(out)
    print('Arquivo salvo.')

# -*- coding: utf-8 -*-
# Patch: habilita/desabilita campos Ad Valorem vs Ad Rem para IBS/CBS
# Adiciona AplicarModoIBSCBS e atualiza AplicarModoIS com cores
# Chama AplicarModoIBSCBS nos eventos corretos

FRM = 'C:/Projeto/Compartilhado/Forms/Produtos_Cadastro.frm'
data = open(FRM, 'rb').read()
data = data.replace(b'\r\n', b'\n').replace(b'\r', b'\n')

COR_ATIVO  = b'&H00C0FFFF&'   # ciano claro
COR_BRANCO = b'&HFFFFFFFF&'   # branco

subs = []

# ─────────────────────────────────────────────────────────────
# 1. Substituir AplicarModoIS adicionando cores + chamar AplicarModoIBSCBS
#    e inserir AplicarModoIBSCBS logo apos
# ─────────────────────────────────────────────────────────────
old_modo_is = (
    b'Private Sub AplicarModoIS()\r\n'
    b'    If cboISClasse.ListIndex < 0 Then Exit Sub\r\n'
    b'    Dim sCod As String\r\n'
    b'    sCod = Left(cboISClasse.Text, 6)\r\n'
    b'    Select Case sCod\r\n'
    b'        Case "900001", "900002", "900003", "900020"  \' Ad Valorem (%)\r\n'
    b'            txtISpIS.Enabled = True:  lblISpIS.Enabled = True\r\n'
    b'            txtISvIS.Enabled = False: lblISvIS.Enabled = False\r\n'
    b'        Case "900010", "900011"  \' Ad Rem (R$/L)\r\n'
    b'            txtISpIS.Enabled = False: lblISpIS.Enabled = False\r\n'
    b'            txtISvIS.Enabled = True:  lblISvIS.Enabled = True\r\n'
    b'        Case "900040"  \' Misto\r\n'
    b'            txtISpIS.Enabled = True:  lblISpIS.Enabled = True\r\n'
    b'            txtISvIS.Enabled = True:  lblISvIS.Enabled = True\r\n'
    b'        Case Else\r\n'
    b'            txtISpIS.Enabled = True:  lblISpIS.Enabled = True\r\n'
    b'            txtISvIS.Enabled = False: lblISvIS.Enabled = False\r\n'
    b'    End Select\r\n'
    b'End Sub'
)

new_modo_is = (
    b'Private Sub AplicarModoIS()\r\n'
    b'    If cboISClasse.ListIndex < 0 Then\r\n'
    b'        txtISpIS.Enabled = False: txtISpIS.BackColor = ' + COR_BRANCO + b'\r\n'
    b'        lblISpIS.Enabled = False\r\n'
    b'        txtISvIS.Enabled = False: txtISvIS.BackColor = ' + COR_BRANCO + b'\r\n'
    b'        lblISvIS.Enabled = False\r\n'
    b'        Exit Sub\r\n'
    b'    End If\r\n'
    b'    Dim sCod As String\r\n'
    b'    sCod = Left(cboISClasse.Text, 6)\r\n'
    b'    Select Case sCod\r\n'
    b'        Case "900001", "900002", "900003", "900020"  \' Ad Valorem (%)\r\n'
    b'            txtISpIS.Enabled = True:  txtISpIS.BackColor = ' + COR_ATIVO + b'\r\n'
    b'            lblISpIS.Enabled = True\r\n'
    b'            txtISvIS.Enabled = False: txtISvIS.BackColor = ' + COR_BRANCO + b'\r\n'
    b'            lblISvIS.Enabled = False\r\n'
    b'        Case "900010", "900011"  \' Ad Rem (R$/L)\r\n'
    b'            txtISpIS.Enabled = False: txtISpIS.BackColor = ' + COR_BRANCO + b'\r\n'
    b'            lblISpIS.Enabled = False\r\n'
    b'            txtISvIS.Enabled = True:  txtISvIS.BackColor = ' + COR_ATIVO + b'\r\n'
    b'            lblISvIS.Enabled = True\r\n'
    b'        Case "900040"  \' Misto\r\n'
    b'            txtISpIS.Enabled = True:  txtISpIS.BackColor = ' + COR_ATIVO + b'\r\n'
    b'            lblISpIS.Enabled = True\r\n'
    b'            txtISvIS.Enabled = True:  txtISvIS.BackColor = ' + COR_ATIVO + b'\r\n'
    b'            lblISvIS.Enabled = True\r\n'
    b'        Case Else\r\n'
    b'            txtISpIS.Enabled = True:  txtISpIS.BackColor = ' + COR_ATIVO + b'\r\n'
    b'            lblISpIS.Enabled = True\r\n'
    b'            txtISvIS.Enabled = False: txtISvIS.BackColor = ' + COR_BRANCO + b'\r\n'
    b'            lblISvIS.Enabled = False\r\n'
    b'    End Select\r\n'
    b'    AplicarModoIBSCBS\r\n'
    b'End Sub\r\n'
    b'\r\n'
    b'Private Sub AplicarModoIBSCBS()\r\n'
    b'    Dim sCls   As String\r\n'
    b'    Dim sCat   As String\r\n'
    b'    Dim sISCls As String\r\n'
    b'    Dim bAdRem As Boolean\r\n'
    b'    Dim bMisto As Boolean\r\n'
    b'\r\n'
    b'    sCls   = Left(cboIBSCBSClasse.Text & "      ", 6)\r\n'
    b'    sCat   = UCase(cboCategoria.Text)\r\n'
    b'    sISCls = Left(cboISClasse.Text & "      ", 6)\r\n'
    b'\r\n'
    b'    bAdRem = False: bMisto = False\r\n'
    b'\r\n'
    b'    \' Regra por cClassTrib IBS/CBS: iniciado em 600 = Ad Rem por UN\r\n'
    b'    If Left(sCls, 3) = "600" Then bAdRem = True\r\n'
    b'\r\n'
    b'    \' Regra por cClassTrib IS\r\n'
    b'    Select Case Trim(sISCls)\r\n'
    b'        Case "900010", "900011": bAdRem = True\r\n'
    b'        Case "900040":           bMisto = True\r\n'
    b'    End Select\r\n'
    b'\r\n'
    b'    \' Regra por categoria\r\n'
    b'    If sCat = "COMBUST\xcdVEIS" Or sCat = "LUBRIFICANTES" Then bAdRem = True\r\n'
    b'\r\n'
    b'    If bMisto Then\r\n'
    b'        \' Misto: ambos ativos\r\n'
    b'        txtCBSpAliq.Enabled   = True:  txtCBSpAliq.BackColor   = ' + COR_ATIVO + b'\r\n'
    b'        txtIBSUFpAliq.Enabled = True:  txtIBSUFpAliq.BackColor = ' + COR_ATIVO + b'\r\n'
    b'        txtIBSMunpAliq.Enabled = True: txtIBSMunpAliq.BackColor = ' + COR_ATIVO + b'\r\n'
    b'        txtCBSvAliq.Enabled   = True:  txtCBSvAliq.BackColor   = ' + COR_ATIVO + b'\r\n'
    b'        txtIBSUFvAliq.Enabled = True:  txtIBSUFvAliq.BackColor = ' + COR_ATIVO + b'\r\n'
    b'        txtIBSMunvAliq.Enabled = True: txtIBSMunvAliq.BackColor = ' + COR_ATIVO + b'\r\n'
    b'    ElseIf bAdRem Then\r\n'
    b'        \' Ad Rem: somente valores monetarios\r\n'
    b'        txtCBSpAliq.Enabled   = False: txtCBSpAliq.BackColor   = ' + COR_BRANCO + b'\r\n'
    b'        txtIBSUFpAliq.Enabled = False: txtIBSUFpAliq.BackColor = ' + COR_BRANCO + b'\r\n'
    b'        txtIBSMunpAliq.Enabled = False: txtIBSMunpAliq.BackColor = ' + COR_BRANCO + b'\r\n'
    b'        txtCBSvAliq.Enabled   = True:  txtCBSvAliq.BackColor   = ' + COR_ATIVO + b'\r\n'
    b'        txtIBSUFvAliq.Enabled = True:  txtIBSUFvAliq.BackColor = ' + COR_ATIVO + b'\r\n'
    b'        txtIBSMunvAliq.Enabled = True: txtIBSMunvAliq.BackColor = ' + COR_ATIVO + b'\r\n'
    b'    Else\r\n'
    b'        \' Ad Valorem: somente percentuais\r\n'
    b'        txtCBSpAliq.Enabled   = True:  txtCBSpAliq.BackColor   = ' + COR_ATIVO + b'\r\n'
    b'        txtIBSUFpAliq.Enabled = True:  txtIBSUFpAliq.BackColor = ' + COR_ATIVO + b'\r\n'
    b'        txtIBSMunpAliq.Enabled = True: txtIBSMunpAliq.BackColor = ' + COR_ATIVO + b'\r\n'
    b'        txtCBSvAliq.Enabled   = False: txtCBSvAliq.BackColor   = ' + COR_BRANCO + b'\r\n'
    b'        txtIBSUFvAliq.Enabled = False: txtIBSUFvAliq.BackColor = ' + COR_BRANCO + b'\r\n'
    b'        txtIBSMunvAliq.Enabled = False: txtIBSMunvAliq.BackColor = ' + COR_BRANCO + b'\r\n'
    b'    End If\r\n'
    b'End Sub'
)
subs.append(('AplicarModoIS + AplicarModoIBSCBS', old_modo_is, new_modo_is))

# ─────────────────────────────────────────────────────────────
# 2. cboISClasse_Click: ja chama AplicarModoIS que por sua vez chama
#    AplicarModoIBSCBS — nenhuma mudanca necessaria
#
# 3. cboIBSCBSClasse_Click: chamar AplicarModoIBSCBS apos rDesc.Close
# ─────────────────────────────────────────────────────────────
old = (b'    If rDesc.State <> 0 Then rDesc.Close\r\n'
       b'End Sub\r\n'
       b'\r\n'
       b'Private Sub cboISCST_Click()')
new  = (b'    If rDesc.State <> 0 Then rDesc.Close\r\n'
        b'    AplicarModoIBSCBS\r\n'
        b'End Sub\r\n'
        b'\r\n'
        b'Private Sub cboISCST_Click()')
subs.append(('cboIBSCBSClasse_Click chama AplicarModoIBSCBS', old, new))

# ─────────────────────────────────────────────────────────────
# 4. cboISCST "99" — desabilitar com cor branca
# ─────────────────────────────────────────────────────────────
old = (b'            txtISpIS.Enabled = False: lblISpIS.Enabled = False\r\n'
       b'            txtISvIS.Enabled = False: lblISvIS.Enabled = False\r\n'
       b'    End Select\r\n'
       b'End Sub\r\n'
       b'\r\n'
       b'Private Sub cboISClasse_Click()')
new  = (b'            txtISpIS.Enabled = False: txtISpIS.BackColor = ' + COR_BRANCO + b': lblISpIS.Enabled = False\r\n'
        b'            txtISvIS.Enabled = False: txtISvIS.BackColor = ' + COR_BRANCO + b': lblISvIS.Enabled = False\r\n'
        b'    End Select\r\n'
        b'End Sub\r\n'
        b'\r\n'
        b'Private Sub cboISClasse_Click()')
subs.append(('cboISCST_Click "99" adiciona BackColor', old, new))

# ─────────────────────────────────────────────────────────────
# Aplicar
# ─────────────────────────────────────────────────────────────
ok = 0
for item in subs:
    name = item[0]; o = item[1]; n = item[2]
    o_n = o.replace(b'\r\n', b'\n')
    d_n = data.replace(b'\r\n', b'\n')
    count = d_n.count(o_n)
    if count == 0: print(f'NENHUMA: {name}'); continue
    if count > 1:  print(f'AMBIGUO ({count}x): {name}'); continue
    data = d_n.replace(o_n, n.replace(b'\r\n', b'\n'))
    print(f'OK: {name}'); ok += 1

data = data.replace(b'\r\n', b'\n').replace(b'\r', b'\n').replace(b'\n', b'\r\n')
open(FRM, 'wb').write(data)
print(f'\nTotal: {ok}/{len(subs)} aplicadas.')

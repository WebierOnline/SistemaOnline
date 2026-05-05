# -*- coding: utf-8 -*-
# Patch: implementa logica completa do Imposto Seletivo (IS) em Produtos_Cadastro.frm
# 1. Preenche cboISClasse em PreencheIBSCBS
# 2. Adiciona cboISCST_Click e cboISClasse_Click
# 3. Corrige LimparCampos: ISCST default = 99
# 4. cmdNovo_Click: seleciona "99" em cboISCST
# 5. AtualizarReforma: reset ISCST de "00" para "99"
# 6. cboIBSCBSCST_Click: todos os "00" viram "99"

FRM = 'C:/Projeto/Compartilhado/Forms/Produtos_Cadastro.frm'
data = open(FRM, 'rb').read()
data = data.replace(b'\r\n', b'\n').replace(b'\r', b'\n')

subs = []

# ─────────────────────────────────────────────────────────────
# 1. Preenche cboISClasse logo apos AddItem do ISCST
# ─────────────────────────────────────────────────────────────
old = (b'    cboISCST.AddItem "99 - Outras Opera\xe7\xf5es"\r\n'
       b'End Sub')
new  = (b'    cboISCST.AddItem "99 - Outras Opera\xe7\xf5es"\r\n'
        b'\r\n'
        b'    \' Imposto Seletivo - Classifica\xe7\xe3o\r\n'
        b'    cboISClasse.AddItem "900001 - Vinhos, Espumantes e Cacha\xe7as"\r\n'
        b'    cboISClasse.AddItem "900002 - Cervejas e Chopes"\r\n'
        b'    cboISClasse.AddItem "900003 - U\xedsque, Gin, Vodka (Destilados)"\r\n'
        b'    cboISClasse.AddItem "900010 - Refrigerantes e Sucos com A\xe7\xfacar"\r\n'
        b'    cboISClasse.AddItem "900011 - Energ\xe9ticos e Bebidas Esportivas"\r\n'
        b'    cboISClasse.AddItem "900020 - Ve\xedculos (Autom\xf3veis poluentes)"\r\n'
        b'    cboISClasse.AddItem "900040 - Cigarros e fumo"\r\n'
        b'End Sub')
subs.append(('preenche cboISClasse', old, new))

# ─────────────────────────────────────────────────────────────
# 2. Insere cboISCST_Click e cboISClasse_Click antes de TirarEspaco
# ─────────────────────────────────────────────────────────────
old = b'Public Function TirarEspaco(ByVal Value As String) As String'
new  = (
    b'Private Sub cboISCST_Click()\r\n'
    b'    Dim sCST As String\r\n'
    b'    If cboISCST.ListIndex < 0 Then Exit Sub\r\n'
    b'    sCST = Left(cboISCST.Text, 2)\r\n'
    b'    Select Case sCST\r\n'
    b'        Case "00"\r\n'
    b'            lblISCST.Caption = "Incid\xeancia na Origem (Ind\xfastria/Importador)"\r\n'
    b'            cboISClasse.Enabled = True\r\n'
    b'            txtISpIS.Text = "0,00": txtISvIS.Text = "0,00"\r\n'
    b'            AplicarModoIS\r\n'
    b'        Case "01"\r\n'
    b'            lblISCST.Caption = "Sa\xedda Tributada (Com\xe9rcio/Varejo)"\r\n'
    b'            cboISClasse.Enabled = True\r\n'
    b'            txtISpIS.Text = "0,00": txtISvIS.Text = "0,00"\r\n'
    b'            AplicarModoIS\r\n'
    b'        Case "99"\r\n'
    b'            lblISCST.Caption = "Outras Opera\xe7\xf5es"\r\n'
    b'            cboISClasse.Enabled = False\r\n'
    b'            cboISClasse.ListIndex = -1\r\n'
    b'            lblISCSTClass.Caption = ""\r\n'
    b'            txtISpIS.Text = "0,00": txtISvIS.Text = "0,00"\r\n'
    b'            txtISpIS.Enabled = False: lblISpIS.Enabled = False\r\n'
    b'            txtISvIS.Enabled = False: lblISvIS.Enabled = False\r\n'
    b'    End Select\r\n'
    b'End Sub\r\n'
    b'\r\n'
    b'Private Sub cboISClasse_Click()\r\n'
    b'    If cboISClasse.ListIndex < 0 Then lblISCSTClass.Caption = "": Exit Sub\r\n'
    b'    Dim sCod As String\r\n'
    b'    sCod = Left(cboISClasse.Text, 6)\r\n'
    b'    lblISCSTClass.Caption = Mid(cboISClasse.Text, 10)\r\n'
    b'    AplicarModoIS\r\n'
    b'End Sub\r\n'
    b'\r\n'
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
    b'End Sub\r\n'
    b'\r\n'
    b'Public Function TirarEspaco(ByVal Value As String) As String'
)
subs.append(('insere cboISCST_Click cboISClasse_Click AplicarModoIS', old, new))

# ─────────────────────────────────────────────────────────────
# 3. LimparCampos: ISCST default = 99 (index 2) em vez de 0
# ─────────────────────────────────────────────────────────────
old = (b'cboISCST.ListIndex = 0\r\n'
       b'txtISpIS.Text = "0,00"')
new  = (b'SelecionarNoCombo cboISCST, "99", True\r\n'
        b'txtISpIS.Text = "0,00"\r\n'
        b'txtISvIS.Text = "0,00"')
subs.append(('limparcampos ISCST=99', old, new))

# ─────────────────────────────────────────────────────────────
# 4. cmdNovo_Click: adiciona selecao de ISCST=99 apos SetFocus
# ─────────────────────────────────────────────────────────────
old = (b'SSTab2.Tab = 0\r\n'
       b'cboUnidMedida.Text = "UN"\r\n'
       b'txtQuant.Text = "0"\r\n'
       b'txtCodBarra.SetFocus\r\n'
       b'End Sub\r\n'
       b'\r\n'
       b'Private Sub cmdRemoverComp_Click()')
new  = (b'SSTab2.Tab = 0\r\n'
        b'cboUnidMedida.Text = "UN"\r\n'
        b'txtQuant.Text = "0"\r\n'
        b'SelecionarNoCombo cboISCST, "99", True\r\n'
        b'txtCodBarra.SetFocus\r\n'
        b'End Sub\r\n'
        b'\r\n'
        b'Private Sub cmdRemoverComp_Click()')
subs.append(('cmdNovo ISCST=99', old, new))

# ─────────────────────────────────────────────────────────────
# 5. AtualizarReforma: reset ISCST "00" -> "99"
# ─────────────────────────────────────────────────────────────
old = (b'    txtISpIS.Text = "0,00"\r\n'
       b'    SelecionarNoCombo cboISCST, "00", True\r\n'
       b'\r\n'
       b'    Select Case sCat')
new  = (b'    txtISpIS.Text = "0,00"\r\n'
        b'    SelecionarNoCombo cboISCST, "99", True\r\n'
        b'\r\n'
        b'    Select Case sCat')
subs.append(('AtualizarReforma reset ISCST=99', old, new))

# ─────────────────────────────────────────────────────────────
# 6. cboIBSCBSCST_Click: todos SelecionarNoCombo cboISCST "00" -> "99"
#    (casos de isencao/diferimento/etc que nao tem IS)
# ─────────────────────────────────────────────────────────────
old = b'            SelecionarNoCombo cboISCST, "00", True'
new  = b'            SelecionarNoCombo cboISCST, "99", True'
subs.append(('cboIBSCBSCST_Click ISCST "00"->"99"', old, new, True))

# ─────────────────────────────────────────────────────────────
# Aplicar
# ─────────────────────────────────────────────────────────────
ok = 0
for item in subs:
    name    = item[0]; o = item[1]; n = item[2]
    all_occ = item[3] if len(item) > 3 else False
    o_n = o.replace(b'\r\n', b'\n')
    d_n = data.replace(b'\r\n', b'\n')
    count = d_n.count(o_n)
    if count == 0:
        print(f'NENHUMA: {name}'); continue
    if not all_occ and count > 1:
        print(f'AMBIGUO ({count}x): {name}'); continue
    data = d_n.replace(o_n, n.replace(b'\r\n', b'\n'))
    print(f'OK ({count}x): {name}'); ok += 1

data = data.replace(b'\r\n', b'\n').replace(b'\r', b'\n').replace(b'\n', b'\r\n')
open(FRM, 'wb').write(data)
print(f'\nTotal: {ok}/{len(subs)} substituicoes aplicadas.')

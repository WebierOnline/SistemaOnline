# -*- coding: utf-8 -*-
# Patch: aplica reducao de IBS/CBS corretamente em Produtos_Cadastro.frm
# Problemas corrigidos:
#   1. cboIBSCBSClasse_Click: calcula e aplica aliquota efetiva (base * (1 - reducao/100))
#   2. cboIBSCBSCST_Click: chama cboIBSCBSClasse_Click no final (apos Select Case setar a base)
#   3. CarregarProduto: remove linhas que sobrescrevem a aliquota ja calculada pelo evento chain

FRM = 'C:/Projeto/Compartilhado/Forms/Produtos_Cadastro.frm'

data = open(FRM, 'rb').read()
data = data.replace(b'\r\n', b'\n').replace(b'\r', b'\n')

subs = []

# ─────────────────────────────────────────────────────────────
# 1. cboIBSCBSClasse_Click: aplicar reducao apos carregar pRedCBS/pRedIBS
# ─────────────────────────────────────────────────────────────
old = (b'    If Not rDesc.EOF Then\r\n'
       b'        lblIBSCBSClass.Caption = rDesc("NomecClassTrib") & ": " & rDesc("DescricaocClassTrib")\r\n'
       b'        txtIBSUFpRedAliq.Text = rDesc("pRedIBS")\r\n'
       b'        txtCBSpRedAliq.Text = rDesc("pRedCBS")\r\n'
       b'    Else\r\n'
       b'        lblIBSCBSClass.Caption = ""\r\n'
       b'    End If\r\n'
       b'    If rDesc.State <> 0 Then rDesc.Close\r\n'
       b'End Sub')
new  = (b'    If Not rDesc.EOF Then\r\n'
        b'        lblIBSCBSClass.Caption = rDesc("NomecClassTrib") & ": " & rDesc("DescricaocClassTrib")\r\n'
        b'        Dim dRedIBS As Double, dRedCBS As Double\r\n'
        b'        dRedCBS = Val(Replace(Replace(CStr(IIf(IsNull(rDesc("pRedCBS")), 0, rDesc("pRedCBS"))), ".", ""), ",", "."))\r\n'
        b'        dRedIBS = Val(Replace(Replace(CStr(IIf(IsNull(rDesc("pRedIBS")), 0, rDesc("pRedIBS"))), ".", ""), ",", "."))\r\n'
        b'        txtCBSpRedAliq.Text   = FormatNumber(dRedCBS, 2)\r\n'
        b'        txtIBSUFpRedAliq.Text = FormatNumber(dRedIBS, 2)\r\n'
        b'        \' Aplica reducao na aliquota base (somente para CSTs nao isentos)\r\n'
        b'        Select Case sCST\r\n'
        b'            Case "400", "410", "510", "550"\r\n'
        b'                \' isencao: manter 0,00 conforme definido pelo CST\r\n'
        b'            Case Else\r\n'
        b'                txtCBSpAliq.Text   = FormatNumber(gCBSpAliq   * (1# - dRedCBS / 100#), 2)\r\n'
        b'                txtIBSUFpAliq.Text = FormatNumber(gIBSUFpAliq * (1# - dRedIBS / 100#), 2)\r\n'
        b'                txtIBSMunpAliq.Text = FormatNumber(gIBSMunpAliq * (1# - dRedIBS / 100#), 2)\r\n'
        b'        End Select\r\n'
        b'    Else\r\n'
        b'        lblIBSCBSClass.Caption = ""\r\n'
        b'    End If\r\n'
        b'    If rDesc.State <> 0 Then rDesc.Close\r\n'
        b'End Sub')
subs.append(('cboIBSCBSClasse_Click reducao', old, new))

# ─────────────────────────────────────────────────────────────
# 2. cboIBSCBSCST_Click: chamar cboIBSCBSClasse_Click apos AtualizarCargaReforma
#    para aplicar reducao sobre a base definida pelo Select Case
# ─────────────────────────────────────────────────────────────
old = (b'    AtualizarCargaReforma\r\n'
       b'End Sub\r\n'
       b'\r\n'
       b'Private Sub cboFabricante_KeyPress')
new  = (b'    AtualizarCargaReforma\r\n'
        b'    \' Aplica reducao da classificacao sobre a base definida acima\r\n'
        b'    If cboIBSCBSClasse.ListIndex >= 0 Then cboIBSCBSClasse_Click\r\n'
        b'End Sub\r\n'
        b'\r\n'
        b'Private Sub cboFabricante_KeyPress')
subs.append(('cboIBSCBSCST_Click chama Classe_Click', old, new))

# ─────────────────────────────────────────────────────────────
# 3. CarregarProduto: remover linhas que sobrescrevem a aliquota calculada
#    (o chain de eventos ja calculou corretamente via cboIBSCBSCST_Click)
# ─────────────────────────────────────────────────────────────
old = (b'    txtCBSpAliq.Text   = FormatNumber(gCBSpAliq, 2)\r\n'
       b'    txtIBSUFpAliq.Text  = FormatNumber(gIBSUFpAliq, 2)\r\n'
       b'    txtIBSMunpAliq.Text = FormatNumber(gIBSMunpAliq, 2)\r\n'
       b'    txtISpIS.Text = FormatNumber(ValidateNull(r("ISpIS")), 2)')
new  = (b'    \' txtCBSpAliq/IBSUFpAliq/IBSMunpAliq ja foram calculados com reducao\r\n'
        b'    \' pelo chain: SelecionarNoCombo cboIBSCBSCST -> cboIBSCBSCST_Click -> cboIBSCBSClasse_Click\r\n'
        b'    txtISpIS.Text = FormatNumber(ValidateNull(r("ISpIS")), 2)')
subs.append(('carregar produto remove override', old, new))

# ─────────────────────────────────────────────────────────────
# Aplicar
# ─────────────────────────────────────────────────────────────
ok = 0
for item in subs:
    name = item[0]; o = item[1]; n = item[2]
    o_n = o.replace(b'\r\n', b'\n')
    d_n = data.replace(b'\r\n', b'\n')
    count = d_n.count(o_n)
    if count == 0:
        print(f'NENHUMA ocorrencia: {name}')
        continue
    if count > 1:
        print(f'AMBIGUO ({count}x): {name}')
        continue
    n_n = n.replace(b'\r\n', b'\n')
    data = d_n.replace(o_n, n_n)
    print(f'OK: {name}')
    ok += 1

data = data.replace(b'\r\n', b'\n').replace(b'\r', b'\n').replace(b'\n', b'\r\n')
open(FRM, 'wb').write(data)
print(f'\nTotal: {ok}/{len(subs)} substituicoes aplicadas.')

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

# 1. PreencheIBSCBS - carrega do banco
text = apply(text,
    'Private Sub PreencheIBSCBS()\r\n'
    '\' CST IBS e CBS (Exemplos baseados na Nota T\xe9cnica)\r\n'
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
    'cboIBSCBSCST.AddItem "820 - Tributa\xe7\xe3o em documento espec\xedfico"\r\n'
    '\r\n'
    '\' CST Imposto Seletivo\r\n'
    'cboISCST.AddItem "00 - N\xe3o Incidente"\r\n'
    'cboISCST.AddItem "01 - Incidente"\r\n'
    'End Sub',

    'Private Sub PreencheIBSCBS()\r\n'
    '    Dim rCst As ADODB.Recordset\r\n'
    '    cboIBSCBSCST.Clear\r\n'
    '    RsOpen rCst, "SELECT CST, DescricaoIBSCBS FROM TbIBSCBS ORDER BY CST"\r\n'
    '    Do While Not rCst.EOF\r\n'
    '        cboIBSCBSCST.AddItem rCst("CST") & " - " & rCst("DescricaoIBSCBS")\r\n'
    '        rCst.MoveNext\r\n'
    '    Loop\r\n'
    '    If rCst.State <> 0 Then rCst.Close\r\n'
    '\r\n'
    '    \' CST Imposto Seletivo\r\n'
    '    cboISCST.AddItem "00 - N\xe3o Incidente"\r\n'
    '    cboISCST.AddItem "01 - Incidente"\r\n'
    'End Sub\r\n'
    '\r\n'
    'Private Sub PreencherClasseTrib(ByVal sCST As String)\r\n'
    '    Dim rCls As ADODB.Recordset\r\n'
    '    cboIBSCBSClasse.Clear\r\n'
    '    If Trim(sCST) = "" Then Exit Sub\r\n'
    '    RsOpen rCls, "SELECT cClassTrib FROM TbIBSCBSClassTrib WHERE CST = \'" & sCST & "\' ORDER BY cClassTrib"\r\n'
    '    Do While Not rCls.EOF\r\n'
    '        cboIBSCBSClasse.AddItem rCls("cClassTrib")\r\n'
    '        rCls.MoveNext\r\n'
    '    Loop\r\n'
    '    If rCls.State <> 0 Then rCls.Close\r\n'
    '    If cboIBSCBSClasse.ListCount > 0 Then cboIBSCBSClasse.ListIndex = 0\r\n'
    'End Sub',
    '1-PreencheIBSCBS-DB')

# 2. cboIBSCBSCST_Click - adiciona chamada a PreencherClasseTrib no inicio
text = apply(text,
    'Private Sub cboIBSCBSCST_Click()\r\n'
    '    Dim sCST As String\r\n'
    '    sCST = Left(cboIBSCBSCST.Text, 3)\r\n',

    'Private Sub cboIBSCBSCST_Click()\r\n'
    '    Dim sCST As String\r\n'
    '    sCST = Left(cboIBSCBSCST.Text, 3)\r\n'
    '    PreencherClasseTrib sCST\r\n',
    '2-cboIBSCBSCST_Click-PreencherClasseTrib')

for e in errors: print(e)
if not errors:
    out = text.encode('windows-1252')
    out = out.replace(b'\r\n', b'\n').replace(b'\r', b'\n').replace(b'\n', b'\r\n')
    open(path, 'wb').write(out)
    print('Arquivo salvo.')

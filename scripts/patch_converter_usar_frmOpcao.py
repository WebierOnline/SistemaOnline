data = open(r'C:\Projeto\OnlineCommerce\Forms\NFe_Completa.frm', 'rb').read()
text = data.decode('windows-1252')

old = (
    '    Dim iMode As Integer\r\n'
    '    iMode = MsgBox("Converter " & nMarcados & " pedidos:" & vbCrLf & vbCrLf & _\r\n'
    '        "SIM = Uma NF por pedido" & vbCrLf & _\r\n'
    '        "N\xc3O = Todos em uma \xfanica NF" & vbCrLf & _\r\n'
    '        "CANCELAR = N\xe3o converter", _\r\n'
    '        vbYesNoCancel + vbQuestion, "Online Commerce")\r\n'
    '    If iMode = vbCancel Then Exit Sub\r\n'
    '    bUnirEmUmaNota = (iMode = vbNo)\r\n'
)

new = (
    '    Dim fOpc As New frmOpcaoConversao\r\n'
    '    fOpc.lblPergunta.Caption = "Converter " & nMarcados & " pedido(s) selecionado(s):"\r\n'
    '    fOpc.Show vbModal\r\n'
    '    Dim iResultado As Integer\r\n'
    '    iResultado = fOpc.Resultado\r\n'
    '    Unload fOpc\r\n'
    '    Set fOpc = Nothing\r\n'
    '    If iResultado = 0 Then Exit Sub\r\n'
    '    bUnirEmUmaNota = (iResultado = 2)\r\n'
)

c = text.count(old)
print('count:', c)
if c == 1:
    text = text.replace(old, new)
    out = text.encode('windows-1252')
    out = out.replace(b'\r\n', b'\n').replace(b'\r', b'\n').replace(b'\n', b'\r\n')
    open(r'C:\Projeto\OnlineCommerce\Forms\NFe_Completa.frm', 'wb').write(out)
    print('OK')
else:
    print('ERRO')

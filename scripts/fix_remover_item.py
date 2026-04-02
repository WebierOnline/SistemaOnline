data = open('OnlineCommerce/Forms/NFe_Completa.frm', 'rb').read()

old = (
    b'    KeyCode = 0\r\n'
    b'    \'If chkDesc.Value = 1 Then\r\n'
    b'    TipoSelecaoConsulta = "0"\r\n'
    b'    cboDescricao.SetFocus\r\n'
    b'    cmdRecalcular_Click\r\n'
    b'Exit Sub\r\n'
    b'\'erro:\r\n'
    b'\'MsgBox "Erro no sistema: " & Err.Number & " - " & Err.Description, vbCritical, "SistemasNFe": Exit Sub\r\n'
    b'End Sub'
)

new = (
    b'    KeyCode = 0\r\n'
    b'    TipoSelecaoConsulta = "0"\r\n'
    b'    cboDescricao.SetFocus\r\n'
    b'    AtualizarTotaisNota\r\n'
    b'Exit Sub\r\n'
    b'End Sub'
)

count = data.count(old)
if count == 1:
    data = data.replace(old, new)
    print('1 OK')
else:
    print(f'ERRO: encontrado {count} vezes')

data = data.replace(b'\r\n', b'\n').replace(b'\r', b'\n').replace(b'\n', b'\r\n')
open('OnlineCommerce/Forms/NFe_Completa.frm', 'wb').write(data)
print('Salvo. Tamanho:', len(data))

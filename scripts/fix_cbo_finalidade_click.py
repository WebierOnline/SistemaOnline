data = open('OnlineCommerce/Forms/NFe_Completa.frm', 'rb').read()

old = (
    b'Private Sub cboFinalidade_Change()\r\n'
    b'    AplicarVisibilidadeGridItens\r\n'
    b'End Sub\r\n'
    b'\r\n'
    b'Private Sub cboFinalidade_LostFocus()\r\n'
)
new = (
    b'Private Sub cboFinalidade_Change()\r\n'
    b'    AplicarVisibilidadeGridItens\r\n'
    b'End Sub\r\n'
    b'\r\n'
    b'Private Sub cboFinalidade_Click()\r\n'
    b'    AplicarVisibilidadeGridItens\r\n'
    b'End Sub\r\n'
    b'\r\n'
    b'Private Sub cboFinalidade_LostFocus()\r\n'
)

c = data.count(old)
if c == 1:
    data = data.replace(old, new)
    print('OK')
else:
    print(f'ERRO: {c} ocorrencias')

data = data.replace(b'\r\n', b'\n').replace(b'\r', b'\n').replace(b'\n', b'\r\n')
open('OnlineCommerce/Forms/NFe_Completa.frm', 'wb').write(data)
print('Salvo. Tamanho:', len(data))

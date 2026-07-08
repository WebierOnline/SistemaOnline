data = open('OnlineCommerce/Forms/NFe_Completa.frm', 'rb').read()

# Fix 1: add CodigoNota to DELETE
old1 = (
    b'    dbData.Execute "DELETE FROM NotaFiscalItens WHERE (CodigoProduto = " & GridNotasItens.TextMatrix(GridNotasItens.Row, 3) & ") AND (ITEM = " & GridNotasItens.TextMatrix(GridNotasItens.Row, 1) & ");"'
)
new1 = (
    b'    dbData.Execute "DELETE FROM NotaFiscalItens WHERE CodigoNota = " & Val(txtCodNota.Text) & " AND (CodigoProduto = " & GridNotasItens.TextMatrix(GridNotasItens.Row, 3) & ") AND (ITEM = " & GridNotasItens.TextMatrix(GridNotasItens.Row, 1) & ");"'
)

# Fix 2: move AtualizarTotaisNota BEFORE SetFocus
old2 = (
    b'    KeyCode = 0\r\n'
    b'    TipoSelecaoConsulta = "0"\r\n'
    b'    cboDescricao.SetFocus\r\n'
    b'    AtualizarTotaisNota\r\n'
    b'Exit Sub\r\n'
    b'End Sub'
)
new2 = (
    b'    KeyCode = 0\r\n'
    b'    TipoSelecaoConsulta = "0"\r\n'
    b'    AtualizarTotaisNota\r\n'
    b'    cboDescricao.SetFocus\r\n'
    b'Exit Sub\r\n'
    b'End Sub'
)

for i, (old, new) in enumerate([(old1, new1), (old2, new2)], 1):
    count = data.count(old)
    if count == 1:
        data = data.replace(old, new)
        print(f'{i} OK')
    else:
        print(f'{i} ERRO: encontrado {count} vezes')

data = data.replace(b'\r\n', b'\n').replace(b'\r', b'\n').replace(b'\n', b'\r\n')
open('OnlineCommerce/Forms/NFe_Completa.frm', 'wb').write(data)
print('Salvo. Tamanho:', len(data))

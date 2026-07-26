data = open('OnlineCommerce/Forms/NFe_Completa.frm', 'rb').read()

old = (
    b'Private Sub cmdCancelar_Click()\r\n'
    b'\'On Error GoTo Err_Cancela\r\n'
    b'\r\n'
    b'cmdNovo.Enabled = True\r\n'
    b'cmdSalvar.Enabled = False\r\n'
    b'cmdCancelar.Enabled = False\r\n'
    b'\r\n'
    b'frmNota.Enabled = False\r\n'
    b'frmDestinatario.Enabled = False\r\n'
    b'Tab_Totais.Enabled = False\r\n'
    b'Tab_Produtos.Enabled = False\r\n'
    b'frmItens.Enabled = False\r\n'
    b'TipoSelecaoConsulta = "0"\r\n'
    b'Tab_Totais.Tab = 0\r\n'
    b'Tab_Produtos.Tab = 0\r\n'
    b'\r\n'
    b'If TbNotas.EditMode <> 0 Then TbNotas.CancelUpdate\r\n'
    b'\r\n'
    b'LimparObjetosNota\r\n'
    b'LimparObjetosDestinatario\r\n'
    b'LimparObjetosProduto\r\n'
    b'LimparObjetosNotaTotais\r\n'
    b'LimparObjestosNotaOutros\r\n'
    b'LimparGridItensNota\r\n'
    b'vTipoEdicaoNFe = ""\r\n'
    b'\r\n'
    b'\'txtInfComple.Text = "EMPRESA ME OU EPP OPTANTE PELO SIMPLES NACIONAL N\xc3O GERA DIREITO A CREDITO FISCAL DE ICMS OU ISS."\r\n'
    b'Exit Sub\r\n'
    b'\r\n'
    b'\'Err_Cancela:\r\n'
    b'\'MsgBox "Erro no sistema: " & Err.Number & " - " & Err.Description, vbCritical, "Online Commerce": Exit Sub\r\n'
    b'End Sub'
)

new = (
    b'Private Sub cmdCancelar_Click()\r\n'
    b'\r\n'
    b'If vTipoEdicaoNFe = "Novo" Then\r\n'
    b"    ' Nota criada pelo cmdNovo mas nao salva -- confirma exclusao\r\n"
    b'    If MsgBox("Deseja cancelar a nota em digita\xe7\xe3o? Os dados ser\xe3o exclu\xeddos.", vbQuestion + vbYesNo, "Online Commerce") <> vbYes Then Exit Sub\r\n'
    b'    If txtCodNota.Text <> "" Then\r\n'
    b'        SQLExecuta "DELETE FROM NotaFiscalItens WHERE CodigoNota = " & Val(txtCodNota.Text)\r\n'
    b'        SQLExecuta "DELETE FROM NotaFiscal WHERE CodigoNota = " & Val(txtCodNota.Text)\r\n'
    b'    End If\r\n'
    b'End If\r\n'
    b'\r\n'
    b'cmdNovo.Enabled = True\r\n'
    b'cmdSalvar.Enabled = False\r\n'
    b'cmdCancelar.Enabled = False\r\n'
    b'\r\n'
    b'frmNota.Enabled = False\r\n'
    b'frmDestinatario.Enabled = False\r\n'
    b'Tab_Totais.Enabled = False\r\n'
    b'Tab_Produtos.Enabled = False\r\n'
    b'frmItens.Enabled = False\r\n'
    b'TipoSelecaoConsulta = "0"\r\n'
    b'Tab_Totais.Tab = 0\r\n'
    b'Tab_Produtos.Tab = 0\r\n'
    b'\r\n'
    b'If TbNotas.EditMode <> 0 Then TbNotas.CancelUpdate\r\n'
    b'\r\n'
    b'LimparObjetosNota\r\n'
    b'LimparObjetosDestinatario\r\n'
    b'LimparObjetosProduto\r\n'
    b'LimparObjetosNotaTotais\r\n'
    b'LimparObjestosNotaOutros\r\n'
    b'LimparGridItensNota\r\n'
    b'vTipoEdicaoNFe = ""\r\n'
    b'\r\n'
    b'End Sub'
)

count = data.count(old)
if count == 1:
    data = data.replace(old, new)
    print('1 OK')
else:
    print(f'ERRO: encontrado {count} vezes')
    idx = data.find(b'Private Sub cmdCancelar_Click')
    end = data.find(b'End Sub', idx) + 7
    print(repr(data[idx:end]))

data = data.replace(b'\r\n', b'\n').replace(b'\r', b'\n').replace(b'\n', b'\r\n')
open('OnlineCommerce/Forms/NFe_Completa.frm', 'wb').write(data)
print('Salvo. Tamanho:', len(data))

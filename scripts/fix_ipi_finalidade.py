data = open('OnlineCommerce/Forms/NFe_Completa.frm', 'rb').read()
results = []

# 1. SELECT: IPIcEnq -> IPICST, reordenar IPI (IPICST, IPIpIPI, IPIvIPI)
old1 = (
    b'       "IPIpIPI, IPIvIPI, IPIcEnq " & _\r\n'
    b'       "FROM NotaFiscalItens WHERE CodigoNota = " & Val(txtCodNota.Text)\r\n'
)
new1 = (
    b'       "IPICST, IPIpIPI, IPIvIPI " & _\r\n'
    b'       "FROM NotaFiscalItens WHERE CodigoNota = " & Val(txtCodNota.Text)\r\n'
)
c = data.count(old1)
results.append(f'1. SELECT IPI: {c}')
if c == 1: data = data.replace(old1, new1)

# 2. FormatarGridItensNota headers IPI
old2 = (
    b'      .TextMatrix(0, 24) = "%IPI"\r\n'
    b'      .TextMatrix(0, 25) = "IPI"\r\n'
    b'      .TextMatrix(0, 26) = "cEnq"\r\n'
)
new2 = (
    b'      .TextMatrix(0, 24) = "CST IPI"\r\n'
    b'      .TextMatrix(0, 25) = "%IPI"\r\n'
    b'      .TextMatrix(0, 26) = "IPI"\r\n'
)
c = data.count(old2)
results.append(f'2. Headers IPI: {c}')
if c == 1: data = data.replace(old2, new2)

# 3. FormatarGridItensNota data rows IPI
old3 = (
    b'            .TextMatrix(.rows - 1, 24) = FormatNumber(rTabela("IPIpIPI"), 2)\r\n'
    b'            .TextMatrix(.rows - 1, 25) = FormatNumber(rTabela("IPIvIPI"), 2)\r\n'
    b'            .TextMatrix(.rows - 1, 26) = rTabela("IPIcEnq")\r\n'
)
new3 = (
    b'            .TextMatrix(.rows - 1, 24) = rTabela("IPICST")\r\n'
    b'            .TextMatrix(.rows - 1, 25) = FormatNumber(rTabela("IPIpIPI"), 2)\r\n'
    b'            .TextMatrix(.rows - 1, 26) = FormatNumber(rTabela("IPIvIPI"), 2)\r\n'
)
c = data.count(old3)
results.append(f'3. Dados IPI: {c}')
if c == 1: data = data.replace(old3, new3)

# 4. Adicionar cboFinalidade_LostFocus apos chkICMSST_Click
old4 = (
    b'Private Sub chkICMSST_Click()\r\n'
    b'   AplicarVisibilidadeGridItens\r\n'
    b'End Sub'
)
new4 = (
    b'Private Sub chkICMSST_Click()\r\n'
    b'   AplicarVisibilidadeGridItens\r\n'
    b'End Sub\r\n'
    b'\r\n'
    b'Private Sub cboFinalidade_LostFocus()\r\n'
    b'   AplicarVisibilidadeGridItens\r\n'
    b'End Sub'
)
c = data.count(old4)
results.append(f'4. cboFinalidade_LostFocus: {c}')
if c == 1: data = data.replace(old4, new4)

data = data.replace(b'\r\n', b'\n').replace(b'\r', b'\n').replace(b'\n', b'\r\n')
open('OnlineCommerce/Forms/NFe_Completa.frm', 'wb').write(data)
for r in results: print(r)
print('Salvo. Tamanho:', len(data))

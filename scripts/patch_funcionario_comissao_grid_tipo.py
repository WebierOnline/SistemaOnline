# -*- coding: utf-8 -*-
"""
Patch Funcionario_Comissao.frm:
Faz o grid principal (sub chameleonButton1_Click, bloco "MONTAR O GRID")
respeitar o cboTipo:
  - cboTipo = "VENDA"     -> comportamento atual (parcelas/pedidos)
  - cboTipo = "SERVICOS"  -> consulta a tabela OS (ordens de servico)

Os campos da query de OS sao aliased para os mesmos nomes que
FormatarGrid ja espera (var_codped, DATA_COMPRA, NUMERO, VALOR_FINAL,
var_StatusPgto, var_FormaPgto, NOME, TIPO_PAGAMENTO, PAGAMENTO), entao
FormatarGrid nao precisa mudar.
"""

PATH = r"C:\projeto\OnlineCommerce\Forms\Funcionario_Comissao.frm"

with open(PATH, "rb") as f:
    raw = f.read()

text = raw.decode("cp1252")
lines = text.split("\r\n")


def find_line(substr, start=0):
    for i in range(start, len(lines)):
        if substr in lines[i]:
            return i
    raise SystemExit(f"ERRO: ancora nao encontrada: {substr!r}")


start = find_line("'MONTAR O GRID")
end = find_line("Set r = dbData.OpenRecordset(sSQL, totalRegistros)", start)
assert lines[start + 1].startswith('sSQL = "SELECT parcelas.COD_PEDIDO'), "bloco MONTAR O GRID nao bate"

original_query_lines = lines[start + 1 : end]  # as 3 linhas do sSQL de VENDA (sem o Set r =)

new_block = []
new_block.append("'MONTAR O GRID")
new_block.append('If cboTipo.Text = "SERVIÇOS" Then')
new_block.append(
    '    sSQL = "SELECT OS.COD_PEDIDO AS var_codped, OS.TIPO_PAGAMENTO, OS.DATA_ENTRADA AS DATA_COMPRA, OS.COD_OS AS NUMERO, OS.TOTAL AS VALOR_FINAL, OS.TIPO_OS AS var_FormaPgto, (CASE WHEN OS.STATUS = 1 THEN \'Pago\' ELSE \'À Pagar\' END) AS var_StatusPgto, OS.COD_FUNCIONARIO, cliente.Nome, OS.COD_CLIENTE, OS.PAGAMENTO " & _'
)
new_block.append(
    '        "FROM OS LEFT JOIN cliente ON OS.COD_CLIENTE = cliente.CODIGO " & _'
)
new_block.append(
    '        "WHERE (OS.COD_FUNCIONARIO = " & txtCodFunc.Text & ") AND (OS.STATUS = 1) AND (MONTH(OS.DATA_ENTRADA) = " & cboMes.ListIndex + 1 & ") AND (YEAR(OS.DATA_ENTRADA) = " & cboAno & ") "'
)
new_block.append("Else")
for l in original_query_lines:
    new_block.append("    " + l)
new_block.append("End If")

lines[start : end] = new_block

out_text = "\r\n".join(lines)
out = out_text.encode("cp1252")
out = out.replace(b"\r\n", b"\n").replace(b"\r", b"\n").replace(b"\n", b"\r\n")

with open(PATH, "wb") as f:
    f.write(out)

print("OK - patch aplicado")
print("bytes originais:", len(raw), "bytes finais:", len(out))

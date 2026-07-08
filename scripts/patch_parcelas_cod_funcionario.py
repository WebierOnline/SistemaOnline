# -*- coding: utf-8 -*-
"""
cmdFinalizar_Click: grava COD_FUNCIONARIO (= txtCodFuncAP.Text) nas parcelas:
- A VISTA: em TODAS as parcelas criadas (1-FORMA e as 2 de 2-FORMAS).
- A PRAZO com entrada: SOMENTE na parcela de entrada (numero=1); as
  parcelas seguintes do loop NAO recebem (pedido explicito do usuario:
  "a entrada (somente ela)").
"""

PATH = r"C:\projeto\OrdemServico\Forms\OS_Recapadora.frm"

with open(PATH, "rb") as f:
    raw = f.read()
text = raw.decode("cp1252")
lines = text.split("\r\n")


def find_line_exact(s, start=0, end=None):
    end = end if end is not None else len(lines)
    for i in range(start, end):
        if lines[i] == s:
            return i
    raise SystemExit(f"ERRO: linha exata nao encontrada: {s!r}")


start = find_line_exact("Private Sub cmdFinalizar_Click()")
end = find_line_exact("End Sub", start)

# ---------------------------------------------------------------
# 1) A PRAZO - parcela de ENTRADA (numero=1)
# ---------------------------------------------------------------
i = find_line_exact(
    '              dbData.Execute "INSERT INTO parcelas (codigo, cod_pedido, cod_os,  numero, data, valor, status, TIPO, DIAS_ATRAZO, JUROS, MULTA, DESCONTO, VALOR_FINAL) VALUES (" & _',
    start, end,
)
lines[i] = lines[i].replace(
    "DESCONTO, VALOR_FINAL) VALUES (",
    "DESCONTO, VALOR_FINAL, COD_FUNCIONARIO) VALUES (",
)
j = i + 2
assert lines[j] == '                 Replace(CCur(txtEntrada.Text), ",", ".") & ", 0, \'OS\', 0, 0, 0, 0, " & Replace(CCur(txtEntrada.Text), ",", ".") & ");"'
lines[j] = lines[j].replace(
    '& ");"',
    '& ", " & txtCodFuncAP.Text & ");"',
)
end = find_line_exact("End Sub", start)

# ---------------------------------------------------------------
# 2) A VISTA - 1 FORMA (unica parcela)
# ---------------------------------------------------------------
i = find_line_exact(
    '                sSQL = "INSERT INTO parcelas (codigo, cod_pedido, cod_os, numero, data, valor, VALOR_FINAL, DIAS_ATRAZO, JUROS, MULTA, DESCONTO, TIPO) VALUES (" & _',
    start, end,
)
lines[i] = lines[i].replace(
    "DESCONTO, TIPO) VALUES (",
    "DESCONTO, TIPO, COD_FUNCIONARIO) VALUES (",
)
j = i + 1
assert "'OS');\"" in lines[j]
lines[j] = lines[j].replace("'OS');\"", "'OS', \" & txtCodFuncAP.Text & \");\"")
end = find_line_exact("End Sub", start)

# ---------------------------------------------------------------
# 3) A VISTA - 2 FORMAS - parcela 1 (entrada)
# ---------------------------------------------------------------
i = find_line_exact(
    '                sSQL = "INSERT INTO parcelas (codigo, cod_pedido, cod_os, numero, data, valor, VALOR_FINAL, DIAS_ATRAZO, JUROS, MULTA, DESCONTO) VALUES (" & _',
    start, end,
)
lines[i] = lines[i].replace(
    "DESCONTO) VALUES (",
    "DESCONTO, COD_FUNCIONARIO) VALUES (",
)
j = i + 1
assert lines[j].rstrip().endswith('0, 0, 0, 0);"')
lines[j] = lines[j].replace('0, 0, 0, 0);"', '0, 0, 0, 0, " & txtCodFuncAP.Text & ");"')
end = find_line_exact("End Sub", start)

# ---------------------------------------------------------------
# 4) A VISTA - 2 FORMAS - parcela 2 (restante)
#    (busca a partir de j+1, pois a string-alvo eh identica a do passo 3
#    - senao encontraria de novo a mesma linha da parcela 1)
# ---------------------------------------------------------------
i = find_line_exact(
    '                sSQL = "INSERT INTO parcelas (codigo, cod_pedido, cod_os, numero, data, valor, VALOR_FINAL, DIAS_ATRAZO, JUROS, MULTA, DESCONTO) VALUES (" & _',
    j + 1, end,
)
lines[i] = lines[i].replace(
    "DESCONTO) VALUES (",
    "DESCONTO, COD_FUNCIONARIO) VALUES (",
)
j = i + 1
assert lines[j].rstrip().endswith('0, 0, 0, 0);"')
lines[j] = lines[j].replace('0, 0, 0, 0);"', '0, 0, 0, 0, " & txtCodFuncAP.Text & ");"')

out_text = "\r\n".join(lines)
out = out_text.encode("cp1252")
with open(PATH, "wb") as f:
    f.write(out)

print("OK - patch aplicado")

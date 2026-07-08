# -*- coding: utf-8 -*-
"""
Patch Funcionario_Comissao.frm:
1) Corrige a referencia quebrada Valor_Comissao1/2/3 -> Valor_ComissaoAV1/2/3
   (coluna renomeada no patch do Funcionario_Cadastro.frm)
2) Implementa faixas (tiered) para comissao A Prazo, usando
   Comissao_Prazo1/2/3 + Valor_ComissaoAP1/2/3 (mesmo padrao de A Vista/Recebidos)
3) Implementa faixas (tiered) para comissao de Servicos, usando
   Comissao_Servico1/2/3 + Valor_ComissaoServ1/2/3, com base em
   OS_Servicos_Auto/OS ligadas aos pedidos do funcionario no periodo
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


# ---------------------------------------------------------------
# 1) Corrige referencia quebrada (Valor_Comissao1/2/3 -> Valor_ComissaoAV1/2/3)
# ---------------------------------------------------------------
i = find_line("SELECT Comissao_Avista1, Comissao_Avista2, Comissao_Avista3, Valor_Comissao1")
assert lines[i].count("Valor_Comissao1, Valor_Comissao2, Valor_Comissao3") == 1, "select avista nao bate"
lines[i] = lines[i].replace(
    "Valor_Comissao1, Valor_Comissao2, Valor_Comissao3",
    "Valor_ComissaoAV1, Valor_ComissaoAV2, Valor_ComissaoAV3",
)

i2 = find_line('If vValorTotalAvista > r("Valor_Comissao1")')
lines[i2] = lines[i2].replace('r("Valor_Comissao1")', 'r("Valor_ComissaoAV1")')

i3 = find_line('If vValorTotalAvista < r("Valor_Comissao3")')
lines[i3] = lines[i3].replace('r("Valor_Comissao3")', 'r("Valor_ComissaoAV3")')

# ---------------------------------------------------------------
# 2) Extrai o FROM/JOIN/WHERE original do bloco A Prazo (reaproveitado
#    sem retyping para preservar os literais acentuados intactos)
# ---------------------------------------------------------------
start_aprazo = find_line("SELECT ISNULL(SUM(parcelas.VALOR_FINAL * funcionario.Comissao_Prazo1 / 100)")
openrs_aprazo = find_line("Set r = dbData.OpenRecordset(sSQL, totalRegistros)", start_aprazo)
sql_lines_aprazo = lines[start_aprazo:openrs_aprazo]
from_join_where_aprazo = sql_lines_aprazo[1:]  # tudo apos a linha do SELECT

end_aprazo = find_line("End If", start_aprazo)  # fecha o bloco lblComAPrazo...End If

new_aprazo = []
new_aprazo.append("Dim vValorTotalAPrazo As Currency")
new_aprazo.append('sSQL = "SELECT ISNULL(SUM(parcelas.VALOR_FINAL), 0) AS varTotalAPrazo " & _')
new_aprazo.extend(from_join_where_aprazo)
new_aprazo.append("Set r = dbData.OpenRecordset(sSQL, totalRegistros)")
new_aprazo.append("")
new_aprazo.append("If Not r.EOF Then")
new_aprazo.append('    vValorTotalAPrazo = FormatNumber(ValidateNull(r("varTotalAPrazo")), 2)')
new_aprazo.append("Else")
new_aprazo.append("    vValorTotalAPrazo = FormatNumber(0, 2)")
new_aprazo.append("End If")
new_aprazo.append("")
new_aprazo.append('sSQL = "SELECT Comissao_Prazo1, Comissao_Prazo2, Comissao_Prazo3, Valor_ComissaoAP1, Valor_ComissaoAP2, Valor_ComissaoAP3 " & _')
new_aprazo.append('       "FROM funcionario " & _')
new_aprazo.append('       "WHERE (CODIGO = " & txtCodFunc.Text & ") "')
new_aprazo.append("Set r = dbData.OpenRecordset(sSQL)")
new_aprazo.append("")
new_aprazo.append("Dim vComissaoAPrazo As Currency")
new_aprazo.append("")
new_aprazo.append("If Not r.EOF Then")
new_aprazo.append('    If vValorTotalAPrazo > r("Valor_ComissaoAP1") Then')
new_aprazo.append('        If vValorTotalAPrazo < r("Valor_ComissaoAP3") Then')
new_aprazo.append('            vComissaoAPrazo = FormatNumber(r("Comissao_Prazo2"), 2)')
new_aprazo.append("        Else")
new_aprazo.append('            vComissaoAPrazo = FormatNumber(r("Comissao_Prazo3"), 2)')
new_aprazo.append("        End If")
new_aprazo.append("    Else")
new_aprazo.append('        vComissaoAPrazo = FormatNumber(r("Comissao_Prazo1"), 2)')
new_aprazo.append("    End If")
new_aprazo.append("Else")
new_aprazo.append("    vComissaoAPrazo = FormatNumber(0, 2)")
new_aprazo.append("End If")
new_aprazo.append("")
new_aprazo.append(
    'sSQL = "SELECT ISNULL(SUM(parcelas.VALOR_FINAL * " & Replace(CDbl(vComissaoAPrazo), ",", ".") & " / 100), 0) AS var_ComAprazo, COUNT(parcelas.CODIGO) AS var_ContParcelas " & _'
)
new_aprazo.extend(from_join_where_aprazo)
new_aprazo.append("Set r = dbData.OpenRecordset(sSQL, totalRegistros)")
new_aprazo.append("")
new_aprazo.append("If Not r.EOF Then")
new_aprazo.append('    lblComAPrazoQtde.Caption = Format(r("var_ContParcelas"), "000")')
new_aprazo.append('    lblComAPrazo.Caption = FormatNumber(r("var_ComAprazo"), 2)')
new_aprazo.append("Else")
new_aprazo.append('    lblComAPrazoQtde.Caption = Format(0, "00")')
new_aprazo.append("    lblComAPrazo.Caption = FormatNumber(0, 2)")
new_aprazo.append("End If")

lines[start_aprazo : end_aprazo + 1] = new_aprazo

# ---------------------------------------------------------------
# 3) Bloco de Servicos (hoje zerado) -> tiered com OS_Servicos_Auto
# ---------------------------------------------------------------
start_serv = find_line('lblComServicosQtde.Caption = Format(0, "000")')
end_serv = find_line("lblComServicos.Caption = FormatNumber(0, 2)", start_serv)
assert end_serv == start_serv + 1, "bloco de servicos nao e mais as duas linhas esperadas"

new_serv = []
new_serv.append("Dim vValorTotalServicos As Currency")
new_serv.append('sSQL = "SELECT ISNULL(SUM(sv.total), 0) AS varTotalServicos " & _')
new_serv.append('       "FROM OS_Servicos_Auto sv INNER JOIN OS ON sv.cod_os = OS.COD_OS INNER JOIN pedidos ON OS.COD_PEDIDO = pedidos.COD_PEDIDO " & _')
new_serv.append(
    '       "WHERE (pedidos.TIPO_PEDIDO = \'VENDA\') AND (pedidos.cancelado = 0) AND (pedidos.COD_FUNCIONARIO = " & txtCodFunc.Text & ") AND (MONTH(pedidos.DATA_COMPRA) = " & cboMes.ListIndex + 1 & ") AND (YEAR(pedidos.DATA_COMPRA) = " & cboAno & ") "'
)
new_serv.append("Set r = dbData.OpenRecordset(sSQL, totalRegistros)")
new_serv.append("")
new_serv.append("If Not r.EOF Then")
new_serv.append('    vValorTotalServicos = FormatNumber(ValidateNull(r("varTotalServicos")), 2)')
new_serv.append("Else")
new_serv.append("    vValorTotalServicos = FormatNumber(0, 2)")
new_serv.append("End If")
new_serv.append("")
new_serv.append('sSQL = "SELECT Comissao_Servico1, Comissao_Servico2, Comissao_Servico3, Valor_ComissaoServ1, Valor_ComissaoServ2, Valor_ComissaoServ3 " & _')
new_serv.append('       "FROM funcionario " & _')
new_serv.append('       "WHERE (CODIGO = " & txtCodFunc.Text & ") "')
new_serv.append("Set r = dbData.OpenRecordset(sSQL)")
new_serv.append("")
new_serv.append("Dim vComissaoServicos As Currency")
new_serv.append("")
new_serv.append("If Not r.EOF Then")
new_serv.append('    If vValorTotalServicos > r("Valor_ComissaoServ1") Then')
new_serv.append('        If vValorTotalServicos < r("Valor_ComissaoServ3") Then')
new_serv.append('            vComissaoServicos = FormatNumber(r("Comissao_Servico2"), 2)')
new_serv.append("        Else")
new_serv.append('            vComissaoServicos = FormatNumber(r("Comissao_Servico3"), 2)')
new_serv.append("        End If")
new_serv.append("    Else")
new_serv.append('        vComissaoServicos = FormatNumber(r("Comissao_Servico1"), 2)')
new_serv.append("    End If")
new_serv.append("Else")
new_serv.append("    vComissaoServicos = FormatNumber(0, 2)")
new_serv.append("End If")
new_serv.append("")
new_serv.append(
    'sSQL = "SELECT ISNULL(SUM(sv.total * " & Replace(CDbl(vComissaoServicos), ",", ".") & " / 100), 0) AS var_ComServicos, COUNT(sv.codigo) AS var_ContServicos " & _'
)
new_serv.append('       "FROM OS_Servicos_Auto sv INNER JOIN OS ON sv.cod_os = OS.COD_OS INNER JOIN pedidos ON OS.COD_PEDIDO = pedidos.COD_PEDIDO " & _')
new_serv.append(
    '       "WHERE (pedidos.TIPO_PEDIDO = \'VENDA\') AND (pedidos.cancelado = 0) AND (pedidos.COD_FUNCIONARIO = " & txtCodFunc.Text & ") AND (MONTH(pedidos.DATA_COMPRA) = " & cboMes.ListIndex + 1 & ") AND (YEAR(pedidos.DATA_COMPRA) = " & cboAno & ") "'
)
new_serv.append("Set r = dbData.OpenRecordset(sSQL, totalRegistros)")
new_serv.append("")
new_serv.append("If Not r.EOF Then")
new_serv.append('    lblComServicosQtde.Caption = Format(r("var_ContServicos"), "000")')
new_serv.append('    lblComServicos.Caption = FormatNumber(r("var_ComServicos"), 2)')
new_serv.append("Else")
new_serv.append('    lblComServicosQtde.Caption = Format(0, "000")')
new_serv.append("    lblComServicos.Caption = FormatNumber(0, 2)")
new_serv.append("End If")

lines[start_serv : end_serv + 1] = new_serv

# ---------------------------------------------------------------
# Grava
# ---------------------------------------------------------------
out_text = "\r\n".join(lines)
out = out_text.encode("cp1252")
out = out.replace(b"\r\n", b"\n").replace(b"\r", b"\n").replace(b"\n", b"\r\n")

with open(PATH, "wb") as f:
    f.write(out)

print("OK - patch aplicado")
print("bytes originais:", len(raw), "bytes finais:", len(out))

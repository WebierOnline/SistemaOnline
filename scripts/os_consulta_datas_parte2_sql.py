# -*- coding: utf-8 -*-
"""
OS_Consulta.frm - parte 2: reescreve o bloco de vTipoOS dentro de
MostrarGrid_OS por completo, de forma unificada:
- adiciona campoData (os.DATA_ENTRADA ou os.DATA_TERMINO conforme
  optDataEntrada/optDataTermino) usado nos filtros DATA/PERIODO/MENSAL
- toda consulta (as 6 variantes: CLIENTE/COD.OS/DATA/PERIODO/MENSAL/
  TODOS) x as 3 secoes de vTipoOS agora sempre traz os.DATA_ENTRADA e
  os.DATA_TERMINO no SELECT, para popular as novas colunas do grid
"""

PATH = r"C:\projeto\OrdemServico\Forms\OS_Consulta.frm"

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


SELECT_COLS = (
    "OS.COD_OS, cliente.Nome, os.DATA_ENTRADA, os.DATA_TERMINO, {extra}, "
    "os.status AS var_status, CASE status_os WHEN 1 THEN 'FECHADO' WHEN 0 THEN 'ABERTO' END AS var_status_os, "
    "os.TIPO_PAGAMENTO, os.PAGAMENTO, os.SUBTOTAL, os.ValorDescReal, os.TOTAL"
)
FROM_CLAUSE = (
    "FROM cliente INNER JOIN OS ON cliente.CODIGO = OS.COD_CLIENTE "
    "INNER JOIN {join} ON OS.COD_OS = {join}.COD_OS"
)


def secao(vtipo_if, join_table, extra_cols, indent="    "):
    select_cols = SELECT_COLS.format(extra=extra_cols)
    from_clause = FROM_CLAUSE.format(join=join_table)
    L = []
    L.append(f"{vtipo_if}")
    L.append(f'{indent}If cboConsultaCriterios.Text = "CLIENTE" Then')
    L.append(f'{indent}   If txtCodClienteLocalizar.Text = "" Then Exit Sub')
    L.append(f'{indent}   sSQL = "SELECT DISTINCT {select_cols} " & _')
    L.append(f'{indent}      "{from_clause} WHERE " & varTIPO_OS & " and (os.cod_cliente = " & txtCodClienteLocalizar.Text & ") " & _')
    L.append(f'{indent}      "ORDER BY " & INDICE')
    L.append(f'{indent}ElseIf cboConsultaCriterios.Text = "CÓD. OS" Then')
    L.append(f'{indent}   If cboLocalizar.Text = "" Then Exit Sub')
    L.append(f'{indent}   sSQL = "SELECT DISTINCT {select_cols} " & _')
    L.append(f'{indent}      "{from_clause} WHERE " & varTIPO_OS & " and (os.cod_os = " & cboLocalizar.Text & ") " & _')
    L.append(f'{indent}      "ORDER BY " & INDICE')
    L.append(f'{indent}ElseIf cboConsultaCriterios.Text = "DATA" Then')
    L.append(f'{indent}   If Not IsDate(mskDataConsulta.Text) Then Exit Sub')
    L.append(f'{indent}   sSQL = "SELECT DISTINCT {select_cols} " & _')
    L.append(f'{indent}      "{from_clause} WHERE " & varTIPO_OS & " and (" & campoData & " >= CONVERT(DATETIME, \'" & Format(mskDataConsulta.Text, ocDATA) & "\', 103)) and (" & campoData & " < DATEADD(day, 1, CONVERT(DATETIME, \'" & Format(mskDataConsulta.Text, ocDATA) & "\', 103))) " & _')
    L.append(f'{indent}      "ORDER BY " & INDICE')
    L.append(f'{indent}ElseIf cboConsultaCriterios.Text = "PERÍODO" Then')
    L.append(f'{indent}   If Not IsDate(mskPeriodoInicio.Text) Or Not IsDate(mskPeriodoFim.Text) Then Exit Sub')
    L.append(f'{indent}   sSQL = "SELECT DISTINCT {select_cols} " & _')
    L.append(f'{indent}      "{from_clause} WHERE " & varTIPO_OS & " and (" & campoData & " >= CONVERT(DATETIME, \'" & Format(mskPeriodoInicio.Text, ocDATA) & "\', 103)) and (" & campoData & " < DATEADD(day, 1, CONVERT(DATETIME, \'" & Format(mskPeriodoFim.Text, ocDATA) & "\', 103))) " & _')
    L.append(f'{indent}      "ORDER BY " & INDICE')
    L.append(f'{indent}ElseIf cboConsultaCriterios.Text = "MENSAL" Then')
    L.append(f'{indent}   If cboMesConsulta.Text = "" Or cboAnoConsulta.Text = "" Then Exit Sub')
    L.append(f'{indent}   sSQL = "SELECT DISTINCT {select_cols} " & _')
    L.append(f'{indent}      "{from_clause} WHERE " & varTIPO_OS & " and (MONTH(" & campoData & ") = " & (cboMesConsulta.ListIndex + 1) & ") and (YEAR(" & campoData & ") = " & cboAnoConsulta.Text & ") " & _')
    L.append(f'{indent}      "ORDER BY " & INDICE')
    L.append(f"{indent}Else")
    L.append(f'{indent}    sSQL = "SELECT DISTINCT {select_cols} " & _')
    L.append(f'{indent}        "{from_clause} " & _')
    L.append(f'{indent}        "WHERE " & varTIPO_OS & " " & SITUACAO & var_STATUS & _')
    L.append(f'{indent}        "ORDER BY " & INDICE')
    L.append(f"{indent}End If")
    return L


novo_bloco = []
novo_bloco += secao(
    'If vTipoOS = "Automóveis" Or vTipoOS = "Motocicletas" Or vTipoOS = "Recapadora" Then',
    "OS_Equipamento_Auto",
    "OS_Equipamento_Auto.fabricante, OS_Equipamento_Auto.ano, OS_Equipamento_Auto.modelo",
)
novo_bloco += secao(
    'ElseIf vTipoOS = "Informática" Or vTipoOS = "Celular" Then',
    "OS_Equipamento",
    "OS_Equipamento.fabricante, OS_Equipamento.equipamento, OS_Equipamento.modelo",
)
novo_bloco += secao(
    'ElseIf vTipoOS = "Comunicação Visual" Then',
    "OS_Equipamento",
    "OS_Equipamento.fabricante, OS_Equipamento.equipamento, OS_Equipamento.modelo",
)

# ---------------------------------------------------------------
# localizar o bloco atual (do primeiro "If vTipoOS" ate o "End If"
# que fecha a secao Comunicacao Visual, logo antes do "Else" final)
# ---------------------------------------------------------------
i_start = find_line_exact('If vTipoOS = "Automóveis" Or vTipoOS = "Motocicletas" Or vTipoOS = "Recapadora" Then')
i_elsefinal = find_line_exact("Else", i_start, i_start + 200)
assert lines[i_elsefinal] == "Else", lines[i_elsefinal]
assert lines[i_elsefinal + 1].strip() == "FormatarGrid_OS Nothing", lines[i_elsefinal + 1]
i_endif_final = i_elsefinal - 1
assert lines[i_endif_final] == "    End If", lines[i_endif_final]

# insere calculo de campoData logo antes do bloco (apos var_STATUS)
i_situacao_block_end = find_line_exact("End If", i_start - 30, i_start)
campo_data_calc = [
    "",
    "'campo de data usado nos filtros DATA/PERÍODO/MENSAL",
    "Dim campoData As String",
    "If optDataEntrada.Value = True Then",
    '   campoData = "os.DATA_ENTRADA"',
    "Else",
    '   campoData = "os.DATA_TERMINO"',
    "End If",
]

# substitui o bloco vTipoOS inteiro (i_start ate i_endif_final, inclusive)
lines[i_start : i_endif_final + 1] = novo_bloco

# insere campo_data_calc antes do (agora deslocado) inicio do bloco
i_start2 = find_line_exact('If vTipoOS = "Automóveis" Or vTipoOS = "Motocicletas" Or vTipoOS = "Recapadora" Then')
lines[i_start2:i_start2] = campo_data_calc

out_text = "\r\n".join(lines)
out_text = out_text.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "\r\n")
with open(PATH, "wb") as f:
    f.write(out_text.encode("cp1252"))

print("OK - parte 2 (SQL unificado) aplicada")

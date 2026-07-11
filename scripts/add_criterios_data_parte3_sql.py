# -*- coding: utf-8 -*-
"""
Parte 3: adiciona os ramos ElseIf DATA/PERIODO/MENSAL em cada uma das
3 secoes de vTipoOS dentro de MostrarGrid_OS (OS_Consulta.frm),
inserindo sempre logo antes do "Else" (bloco TODOS) de cada secao.
Processa de baixo pra cima pra nao precisar recalcular indices.
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


def find_bare_else(start, end):
    for i in range(start, end):
        if lines[i] == "    Else":
            return i
    raise SystemExit(f"ERRO: 'Else' nao encontrado entre {start} e {end}")


i_auto = find_line_exact('If vTipoOS = "Automóveis" Or vTipoOS = "Motocicletas" Or vTipoOS = "Recapadora" Then')
i_info = find_line_exact('ElseIf vTipoOS = "Informática" Or vTipoOS = "Celular" Then')
i_comu = find_line_exact('ElseIf vTipoOS = "Comunicação Visual" Then')
i_fim = find_line_exact("Else", i_comu, i_comu + 60)  # bloco final (FormatarGrid_OS Nothing)
# o "Else" final tem 0 de indentacao (nivel do If externo), diferente do "    Else" interno
assert lines[i_fim] == "Else", lines[i_fim]

else_comu = find_bare_else(i_comu, i_fim)
else_info = find_bare_else(i_info, i_comu)
else_auto = find_bare_else(i_auto, i_info)


def bloco_data(join_table, campos_extra):
    return [
        '    ElseIf cboConsultaCriterios.Text = "DATA" Then',
        "       If Not IsDate(mskDataConsulta.Text) Then Exit Sub",
        f"       sSQL = \"SELECT DISTINCT OS.COD_OS, cliente.Nome, {campos_extra}, os.status AS var_status, CASE status_os WHEN 1 THEN 'FECHADO' WHEN 0 THEN 'ABERTO' END AS var_status_os, os.TIPO_PAGAMENTO, os.PAGAMENTO, os.SUBTOTAL, os.ValorDescReal, os.TOTAL \" & _",
        f'          "FROM cliente INNER JOIN OS ON cliente.CODIGO = OS.COD_CLIENTE INNER JOIN {join_table} ON OS.COD_OS = {join_table}.COD_OS WHERE " & varTIPO_OS & " and (os.DATA_ENTRADA >= CONVERT(DATETIME, \'" & Format(mskDataConsulta.Text, ocDATA) & "\', 103)) and (os.DATA_ENTRADA < DATEADD(day, 1, CONVERT(DATETIME, \'" & Format(mskDataConsulta.Text, ocDATA) & "\', 103))) " & _',
        '          "ORDER BY " & INDICE',
        '    ElseIf cboConsultaCriterios.Text = "PERÍODO" Then',
        "       If Not IsDate(mskPeriodoInicio.Text) Or Not IsDate(mskPeriodoFim.Text) Then Exit Sub",
        f"       sSQL = \"SELECT DISTINCT OS.COD_OS, cliente.Nome, {campos_extra}, os.status AS var_status, CASE status_os WHEN 1 THEN 'FECHADO' WHEN 0 THEN 'ABERTO' END AS var_status_os, os.TIPO_PAGAMENTO, os.PAGAMENTO, os.SUBTOTAL, os.ValorDescReal, os.TOTAL \" & _",
        f'          "FROM cliente INNER JOIN OS ON cliente.CODIGO = OS.COD_CLIENTE INNER JOIN {join_table} ON OS.COD_OS = {join_table}.COD_OS WHERE " & varTIPO_OS & " and (os.DATA_ENTRADA >= CONVERT(DATETIME, \'" & Format(mskPeriodoInicio.Text, ocDATA) & "\', 103)) and (os.DATA_ENTRADA < DATEADD(day, 1, CONVERT(DATETIME, \'" & Format(mskPeriodoFim.Text, ocDATA) & "\', 103))) " & _',
        '          "ORDER BY " & INDICE',
        '    ElseIf cboConsultaCriterios.Text = "MENSAL" Then',
        '       If cboMesConsulta.Text = "" Or cboAnoConsulta.Text = "" Then Exit Sub',
        f"       sSQL = \"SELECT DISTINCT OS.COD_OS, cliente.Nome, {campos_extra}, os.status AS var_status, CASE status_os WHEN 1 THEN 'FECHADO' WHEN 0 THEN 'ABERTO' END AS var_status_os, os.TIPO_PAGAMENTO, os.PAGAMENTO, os.SUBTOTAL, os.ValorDescReal, os.TOTAL \" & _",
        f'          "FROM cliente INNER JOIN OS ON cliente.CODIGO = OS.COD_CLIENTE INNER JOIN {join_table} ON OS.COD_OS = {join_table}.COD_OS WHERE " & varTIPO_OS & " and (MONTH(os.DATA_ENTRADA) = " & (cboMesConsulta.ListIndex + 1) & ") and (YEAR(os.DATA_ENTRADA) = " & cboAnoConsulta.Text & ") " & _',
        '          "ORDER BY " & INDICE',
    ]


# secao Comunicacao Visual (OS_Equipamento) - processa primeiro (mais embaixo)
lines[else_comu:else_comu] = bloco_data(
    "OS_Equipamento",
    "OS_Equipamento.fabricante, OS_Equipamento.equipamento, OS_Equipamento.modelo",
)

# secao Informatica/Celular (OS_Equipamento)
lines[else_info:else_info] = bloco_data(
    "OS_Equipamento",
    "OS_Equipamento.fabricante, OS_Equipamento.equipamento, OS_Equipamento.modelo",
)

# secao Automoveis/Motocicletas/Recapadora (OS_Equipamento_Auto)
lines[else_auto:else_auto] = bloco_data(
    "OS_Equipamento_Auto",
    "OS_Equipamento_Auto.fabricante, OS_Equipamento_Auto.ano, OS_Equipamento_Auto.modelo",
)

out_text = "\r\n".join(lines)
out_text = out_text.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "\r\n")
with open(PATH, "wb") as f:
    f.write(out_text.encode("cp1252"))

print("OK - parte 3 (ramos SQL) aplicada")
